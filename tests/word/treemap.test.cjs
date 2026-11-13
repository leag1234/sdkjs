// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Treemap chart (extended chart): the levels, one colour per top-level branch, headers and leaf labels like Office.
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}
const drawer = read('common/Charts/ChartsDrawer.js'), chartSpace = read('common/Drawings/Format/ChartSpace.js');

function load() {
	const ctx = vm.createContext({console});
	vm.runInContext(slice(drawer, 'function squarifyTreemap(', 'function CCachedTreemap(') + ';this.squarify = squarifyTreemap; this.build = buildHierarchicalTreemap;', ctx);
	return ctx;
}
const {squarify, build} = load();
const plain = v => JSON.parse(JSON.stringify(v));
const lvl = values => ({ptCount: values.length, pts: values.map((v, i) => ({idx: i, val: v})).filter(p => p.val !== '')});

// the data of chartex.xlsx made by Excel (Branch / Stem / Leaf / Value), levels root first
const LEVELS = [
	lvl(['Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 3', 'Branch 3', 'Branch 3']),
	lvl(['Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', '', 'Stem 3', 'Stem 3', '', 'Stem 4', 'Stem 4', 'Stem 5']),
	lvl(['Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6', 'Leaf 7', 'Leaf 8', 'Leaf 9', 'Leaf 10', 'Leaf 11'])
];
const VALUES = [22, 12, 18, 87, 30, 88, 25, 43, 52, 33, 64].map((v, i) => ({idx: i, val: v}));
const W = 508, H = 247;

test('squarified rectangles fill the area, in the order of the values, without overlap', () => {
	const rects = squarify([60, 30, 20, 10, 5], 0, 0, 100, 50);
	assert.equal(rects.length, 5);
	const area = rects.reduce((s, r) => s + r.w * r.h, 0);
	assert.ok(Math.abs(area - 5000) < 1e-6);
	for (const r of rects) assert.ok(r.x >= -1e-9 && r.y >= -1e-9 && r.x + r.w <= 100 + 1e-9 && r.y + r.h <= 50 + 1e-9);
	for (let i = 0; i < rects.length; ++i) {
		assert.ok(Math.abs(rects[i].w * rects[i].h - [60, 30, 20, 10, 5][i] / 125 * 5000) < 1e-6);
		for (let j = i + 1; j < rects.length; ++j) {
			const a = rects[i], b = rects[j];
			assert.ok(a.x + a.w <= b.x + 1e-9 || b.x + b.w <= a.x + 1e-9 || a.y + a.h <= b.y + 1e-9 || b.y + b.h <= a.y + 1e-9, 'overlap ' + i + ' ' + j);
		}
	}
	assert.deepEqual(plain(squarify([1], 5, 6, 0, 10)), [{x: 5, y: 6, w: 0, h: 0}]);
});

test('the tree: branches in the order of the data, the leaf of a row without stem under its branch', () => {
	const t = build(VALUES, LEVELS, W, H);
	assert.deepEqual(plain(t.branches.map(b => b.name)), ['Branch 1', 'Branch 2', 'Branch 3']);
	assert.deepEqual(plain(t.branches.map(b => b.leaves.slice().sort((a, b) => a - b))), [[0, 1, 2, 3, 4], [5, 6, 7], [8, 9, 10]]);
	assert.equal(t.rects.length, 11);
	assert.equal(t.leafNames[4], 'Leaf 5', 'a leaf without a stem keeps its name');
	assert.equal(t.leafNames[0], 'Leaf 1');
});

test('each branch is one area, in proportion to its sum, and its leaves tile it', () => {
	const t = build(VALUES, LEVELS, W, H);
	const sums = [169, 156, 149], total = 474;
	t.branches.forEach((b, k) => {
		const area = t.rects.filter(r => r.branch === k).reduce((s, r) => s + r.w * r.h, 0);
		assert.ok(Math.abs(area - W * H * sums[k] / total) < 1e-6, 'area of branch ' + k);
	});
	for (const r of t.rects) {
		const v = VALUES[r.idx].val;
		assert.ok(Math.abs(r.w * r.h - W * H * v / total) < 1e-6, 'area of leaf ' + r.idx);
		assert.ok(r.x >= -1e-9 && r.y >= -1e-9 && r.x + r.w <= W + 1e-9 && r.y + r.h <= H + 1e-9);
	}
	// the three branches of a wide chart stand side by side, as in Office; the branch of the first rows is the left one
	const left = k => Math.min(...t.rects.filter(r => r.branch === k).map(r => r.x));
	assert.ok(left(0) < left(1) && left(1) < left(2));
});

test('a stem is a group inside its branch: its leaves share one rectangle', () => {
	const t = build(VALUES, LEVELS, W, H);
	const box = ids => {
		const rs = t.rects.filter(r => ids.includes(r.idx));
		const x0 = Math.min(...rs.map(r => r.x)), y0 = Math.min(...rs.map(r => r.y));
		const x1 = Math.max(...rs.map(r => r.x + r.w)), y1 = Math.max(...rs.map(r => r.y + r.h));
		return {area: (x1 - x0) * (y1 - y0), sum: rs.reduce((s, r) => s + r.w * r.h, 0)};
	};
	for (const ids of [[0, 1], [2, 3], [5, 6], [8, 9]]) {
		const b = box(ids);
		assert.ok(Math.abs(b.area - b.sum) < 1e-6, 'the leaves of ' + ids + ' form a rectangle');
	}
});

test('rows with no value, no name or a negative value get no rectangle; a flat list is one branch per row', () => {
	const t = build([{idx: 0, val: 10}, {idx: 1, val: 0}, {idx: 2, val: -4}, {idx: 3, val: 5}],
		[lvl(['A', 'A', 'A', 'B']), lvl(['a1', 'a2', 'a3', 'b1'])], 100, 100);
	assert.deepEqual(plain(t.rects.map(r => r.idx).sort()), [0, 3]);
	assert.deepEqual(plain(t.branches.map(b => b.name)), ['A', 'B']);
	const none = build([], [lvl(['A'])], 10, 10);
	assert.deepEqual(plain(none.rects), []);
});

// ------------------------------------------------------------------------------ the drawer
function loadDrawer() {
	const AscFormat = {isRealNumber: v => typeof v === 'number' && isFinite(v)};
	const ctx = vm.createContext({AscFormat, console});
	vm.runInContext(slice(drawer, 'function drawTreemapChart(', '\n/** @constructor */\nfunction ') + ';this.d = drawTreemapChart;', ctx);
	return ctx;
}
function makeChart(t) {
	const ctx = loadDrawer();
	const paths = [];
	const chartSpaceStub = {
		getCachedData: () => ({data: [], rects: t.rects, branches: t.branches}), chart: {plotArea: {x: 10, y: 20}},
		getBranchBrush: k => ({accent: k % 6}), GetPath: id => paths[id]
	};
	const chart = new ctx.d(null, {
		calcProp: {chartGutter: {_left: 0, _top: 0}, pxToMM: 1, trueWidth: W, trueHeight: H},
		cChartSpace: chartSpaceStub,
		_calculateRect: (x, y, w, h) => { paths.push({cmds: [{X: x, Y: y}, {X: x, Y: y - h}, {X: x + w, Y: y - h}]}); return paths.length - 1; }
	});
	paths.forEach(() => {});
	for (const p of paths) {}
	return {chart, paths, chartSpaceStub};
}
function withPaths(chart, paths) {
	chart.cChartSpace.GetPath = id => ({getCommandByIndex: i => paths[id].cmds[i]});
	return chart;
}
test('a leaf is drawn in the colour of its branch, an explicit data point fill of the file stays', () => {
	const t = build(VALUES, LEVELS, W, H);
	const {chart, paths} = makeChart(t);
	chart.recalculate();
	withPaths(chart, paths);
	const series = {getDptByIdx: () => null};
	assert.equal(chart._getBranchBrush(series, 0).accent, 0);
	assert.equal(chart._getBranchBrush(series, 4).accent, 0, 'Leaf 5 has no stem: still Branch 1');
	assert.equal(chart._getBranchBrush(series, 6).accent, 1);
	assert.equal(chart._getBranchBrush(series, 10).accent, 2);
	assert.equal(chart._getBranchBrush({getDptByIdx: i => i === 3 ? {spPr: {Fill: {}}} : null}, 3), null);
	assert.equal(chart._getBranchBrush(series, 999), null);
});
test('the header of a branch is at the top left of the leaf in the corner of the branch', () => {
	const t = build(VALUES, LEVELS, W, H);
	const {chart, paths} = makeChart(t);
	chart.recalculate();
	withPaths(chart, paths);
	for (const [k, name] of [[0, 'Branch 1'], [1, 'Branch 2'], [2, 'Branch 3']]) {
		const corner = t.rects.filter(r => r.branch === k).reduce((a, r) => (r.y < a.y - 1e-9 || (Math.abs(r.y - a.y) < 1e-9 && r.x < a.x)) ? r : a);
		const lbl = {eoBranch: k, idx: -1 - k, extX: 20, extY: 8};
		const pos = chart._calculateDLbl(lbl);
		assert.equal(lbl.eoHidden, false, name);
		assert.ok(Math.abs(pos.x - (10 + corner.x + 2)) < 1e-6, name + ' x');
		assert.ok(Math.abs(pos.y - (20 + corner.y + 2)) < 1e-6, name + ' y');
	}
	const huge = {eoBranch: 0, idx: -1, extX: 1000, extY: 8};
	chart._calculateDLbl(huge);
	assert.equal(huge.eoHidden, true, 'a header wider than its leaf is hidden');
});
test('a leaf label is at the bottom left of its leaf and hidden when it does not fit', () => {
	const t = build(VALUES, LEVELS, W, H);
	const {chart, paths} = makeChart(t);
	chart.recalculate();
	withPaths(chart, paths);
	const r = t.rects.find(x => x.idx === 3);
	const lbl = {idx: 3, extX: 20, extY: 8};
	const pos = chart._calculateDLbl(lbl);
	assert.equal(lbl.eoHidden, false);
	assert.ok(Math.abs(pos.x - (10 + r.x + 2)) < 1e-6);
	assert.ok(Math.abs(pos.y - (20 + r.y + r.h - 8 - 2)) < 1e-6);
	const wide = {idx: 3, extX: r.w + 5, extY: 8};
	chart._calculateDLbl(wide);
	assert.equal(wide.eoHidden, true);
});

// ------------------------------------------------------------------------------ the wiring
test('the chart space: leaf names, headers, branch colours and legend entries', () => {
	for (const needle of ['cachedData.leafNames && cachedData.leafNames[pt.idx]', 'calculateTreemapHeaders = function', 'CChartSpace.prototype.getBranchBrush = function',
		'getHierarchyLegendEntries = function', 'AscFormat.PARENT_LABEL_LAYOUT_NONE', 'oCachedData.type === AscFormat.SERIES_LAYOUT_TREEMAP'])
		assert(chartSpace.includes(needle), needle);
	assert(chartSpace.indexOf('const aBranchEntries = this.getHierarchyLegendEntries()') > 0);
});
test('the legend entries are the top-level branches in the order of the data, read from the series', () => {
	const src = slice(chartSpace, 'CChartSpace.prototype.getHierarchyLegendEntries = function', 'CChartSpace.prototype.calculateSunburstDLbls');
	const ctx = vm.createContext({AscFormat: {SERIES_LAYOUT_SUNBURST: 5, SERIES_LAYOUT_TREEMAP: 4}});
	vm.runInContext('var CChartSpace = function () {};' + src + ';this.f = CChartSpace.prototype.getHierarchyLegendEntries;', ctx);
	const mk = (layoutId, levels, nums) => ({
		isChartEx: () => true,
		chart: {plotArea: {plotAreaRegion: {series: [{layoutId, getCatLit: () => levels, getValLit: () => ({pts: nums})}]}}}
	});
	assert.deepEqual(plain(ctx.f.call(mk(4, LEVELS, VALUES))), [{name: 'Branch 1'}, {name: 'Branch 2'}, {name: 'Branch 3'}]);
	assert.deepEqual(plain(ctx.f.call(mk(5, LEVELS, VALUES))), [{name: 'Branch 1'}, {name: 'Branch 2'}, {name: 'Branch 3'}]);
	assert.equal(ctx.f.call(mk(4, [LEVELS[2]], VALUES)), null, 'one level: no branches');
	assert.equal(ctx.f.call(mk(2, LEVELS, VALUES)), null, 'a waterfall has none');
	assert.equal(ctx.f.call({isChartEx: () => false}), null);
	const names = ctx.f.call(mk(4, LEVELS, VALUES.map((v, i) => ({idx: i, val: i < 5 ? 0 : v.val}))));
	assert.deepEqual(plain(names), [{name: 'Branch 2'}, {name: 'Branch 3'}], 'a branch with no value has no entry');
});
