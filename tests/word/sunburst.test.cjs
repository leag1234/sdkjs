// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Sunburst chart (extended chart, cx:chartSpace): level order, colours per branch, labels and the label drag.
// The code under test is cut out of the sources and run in a vm with small stubs, like the other tests of this folder.
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
const plain = v => JSON.parse(JSON.stringify(v));
const drawer = read('common/Charts/ChartsDrawer.js'), common = read('common/editorscommon.js'),
	chartEx = read('common/Drawings/Format/ChartEx.js'), format = read('common/Drawings/Format/ChartFormat.js'),
	moves = read('common/Drawings/TrackObjects/MoveTracks.js');

// ------------------------------------------------------------------------------ the tree of the chart
function loadTree() {
	const ctx = vm.createContext({console});
	const treeSrc = slice(common, 'function CTree() {', '\n\tfunction ');
	ctx.AscCommon = {};
	ctx.AscFormat = {InitClassWithoutType(child, parent) { child.prototype = Object.create(parent.prototype); child.prototype.constructor = child; }};
	ctx.CCachedChartExData = function () {};
	vm.runInContext(treeSrc + ';AscCommon.CTree = CTree;', ctx);
	vm.runInContext(slice(drawer, 'function CSunburstNode(', 'function CUpDownBars(') + ';this.CCachedSunburst = CCachedSunburst;', ctx);
	return ctx;
}
// the file has points with no text (<cx:pt idx="5"></cx:pt>): they are kept here, as the reader keeps them
const lvl = (values) => ({ptCount: values.length, pts: values.map((v, i) => ({idx: i, val: v}))});
// chartex.docx of the issue #76: the levels are stored from the deepest (leaf) to the shallowest (branch)
const FILE_LEVELS = [
	lvl(['Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', '', '', 'Leaf 8', '', 'Leaf 10', 'Leaf 11', 'Leaf 12', 'Leaf 13', 'Leaf 14', 'Leaf 15', '']),
	lvl(['Stem 1', 'Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', 'Leaf 6', 'Leaf 7', 'Stem 3', 'Leaf 9', 'Stem 4', 'Stem 4', 'Stem 5', 'Stem 5', 'Stem 6', 'Stem 6', 'Leaf 16']),
	lvl(['Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3'])
];
const SIZES = [22, 12, 18, 87, 88, 17, 14, 25, 16, 24, 89, 16, 19, 86, 23, 21].map((v, i) => ({idx: i, val: v}));

// CSeries.getCatLit
function loadCatLit() {
	const ctx = vm.createContext({AscFormat: {SERIES_LAYOUT_SUNBURST: 5, SERIES_LAYOUT_TREEMAP: 4, SERIES_LAYOUT_WATERFALL: 2, SERIES_LAYOUT_FUNNEL: 3}});
	vm.runInContext('var CSeries = function () {};' + slice(chartEx, 'CSeries.prototype.getCatLit = function', 'CSeries.prototype.getBrush') + ';this.getCatLit = CSeries.prototype.getCatLit;', ctx);
	return ctx.getCatLit;
}
const seriesWith = levels => ({getData: () => ({getCatDimensions: () => [{levelData: levels}]})});

function names(node) {return node.getChildren().map(c => c.name).sort();}
function build(levelsForTree) {
	const ctx = loadTree();
	const cached = new ctx.CCachedSunburst(0, {Id: 'x'}, {pts: SIZES}, levelsForTree);
	return cached.data;
}
test('the last cx:lvl is the root of the tree: Branch, Stem, Leaf', () => {
	const getCatLit = loadCatLit();
	const levels = getCatLit.call(seriesWith(FILE_LEVELS), 5);
	assert.equal(levels[0].pts[0].val, 'Branch 1');
	assert.equal(levels[2].pts[0].val, 'Leaf 1');
	assert.equal(FILE_LEVELS[0].pts[0].val, 'Leaf 1', 'the model keeps the order of the file');
	const tree = build(levels);
	assert.deepEqual(plain(names(tree)), ['Branch 1', 'Branch 2', 'Branch 3']);
	const byName = Object.fromEntries(tree.getChildren().map(c => [c.name, c]));
	assert.deepEqual(plain(names(byName['Branch 1'])), ['Leaf 6', 'Leaf 7', 'Stem 1', 'Stem 2']);
	assert.deepEqual(plain(names(byName['Branch 2'])), ['Leaf 9', 'Stem 3', 'Stem 4']);
	assert.deepEqual(plain(names(byName['Branch 3'])), ['Leaf 16', 'Stem 5', 'Stem 6']);
	const stem1 = byName['Branch 1'].getChildren().find(c => c.name === 'Stem 1');
	assert.deepEqual(plain(names(stem1)), ['Leaf 1', 'Leaf 2', 'Leaf 3']);
});
test('a leaf without a stem stays in the middle ring, its outer ring is empty', () => {
	const tree = build(loadCatLit().call(seriesWith(FILE_LEVELS), 5));
	const branch1 = tree.getChildren().find(c => c.name === 'Branch 1');
	for (const name of ['Leaf 6', 'Leaf 7']) {
		assert.equal(branch1.getChildren().find(c => c.name === name).children.length, 0, name);
	}
	const middle = tree.getChildren().reduce((all, b) => all.concat(b.getChildren()), []);
	for (const name of ['Leaf 9', 'Leaf 16']) {
		assert.equal(middle.find(c => c.name === name).children.length, 0, name);
	}
});
test('segments per ring: 3 branches, 10 middle (6 stems and 4 lone leaves), 12 outer', () => {
	const tree = build(loadCatLit().call(seriesWith(FILE_LEVELS), 5));
	const rings = [[], [], [], []];
	const walk = (node, depth) => node.getChildren().forEach(c => { rings[depth].push(c); walk(c, depth + 1); });
	walk(tree, 0);
	assert.deepEqual(rings.map(r => r.length), [3, 10, 12, 0]);
	assert.ok(rings.every(r => r.every(n => n.name !== '' && n.name !== null)), 'no nameless segment');
});
test('the labels of the other chart types come from the deepest level, the treemap reads its levels like the sunburst', () => {
	const getCatLit = loadCatLit();
	assert.equal(getCatLit.call(seriesWith(FILE_LEVELS), 99), FILE_LEVELS[0]);          // clustered column, box and whisker: the first level
	assert.equal(getCatLit.call(seriesWith(FILE_LEVELS), 2), FILE_LEVELS[0]);           // waterfall: the leaves, like Office
	assert.equal(getCatLit.call(seriesWith(FILE_LEVELS), 3), FILE_LEVELS[0]);           // funnel
	const treemap = getCatLit.call(seriesWith(FILE_LEVELS), 4);
	assert.equal(treemap[0].pts[0].val, 'Branch 1');
	assert.equal(treemap[2].pts[0].val, 'Leaf 1');
});
test('a sheet gives the levels root first: the model keeps them deepest first, as the file does', () => {
	const src = slice(chartEx, 'CDimension.prototype.updateReferences = function', 'CDimension.prototype.updateCache');
	assert(src.includes('aNewLevels.reverse()') && src.includes('this.type === AscFormat.STRING_DIMENSION_TYPE_CAT'));
	assert(!chartEx.includes('isSunburstCategories'));
});

// ------------------------------------------------------------------------------ colours and labels
function loadDrawer() {
	const calc = [];
	const AscFormat = {
		isRealNumber: v => typeof v === 'number' && isFinite(v),
		CreateUnifillSolidFillSchemeColor: (id) => ({accent: id, calculate() { calc.push(id); }}),
		ExecuteNoHistory: (f, t, a) => f.apply(t, a || [])
	};
	const ctx = vm.createContext({AscFormat, cToDeg: 180 / Math.PI, console});
	ctx.AscCommon = {};
	vm.runInContext(slice(common, 'function CTree() {', '\n\tfunction ') + ';AscCommon.CTree = CTree;', ctx);
	vm.runInContext(slice(drawer, 'function drawSunburstChart(', '/** @constructor */\nfunction drawWaterfallChart(') + ';this.drawSunburstChart = drawSunburstChart;', ctx);
	return {ctx, calc};
}
function makeChart(tree, layers) {
	const {ctx, calc} = loadDrawer();
	const chartSpace = {
		getCachedData: () => ({data: tree, layersCount: layers}), getParentObjects: () => ({}), clrMapOvr: null,
		AllocPath: () => 1, GetPath: () => ({moveTo() {}, arcTo() {}, lnTo() {}})
	};
	const props = {trueWidth: 100, trueHeight: 100, chartGutter: {_left: 0, _top: 0}, pxToMM: 1, pathW: 1, pathH: 1};
	const chart = new ctx.drawSunburstChart(null, {calcProp: props, cChartSpace: chartSpace, cChartDrawer: {}});
	return {chart, calc};
}
function treeOf(spec) {
	const ctx = loadTree();
	const mk = (s, parent) => {
		const n = new ctx.CSunburstNode(s.name, s.v, s.i, parent, [], s.st, 0);
		n.children = (s.kids || []).map(k => mk(k, n));
		return n;
	};
	const head = new ctx.CSunburstNode(null, 1, -1, null, [], 0, 0);
	head.children = spec.map(s => mk(s, head));
	return head;
}
// ring sorted by size (ascending), the data order is the order of stIdx
const SPEC = [
	{name: 'B2', v: 0.2, i: 0, st: 7, kids: [{name: 'S3', v: 0.2, i: 1, st: 7}]},
	{name: 'B3', v: 0.3, i: 2, st: 11},
	{name: 'B1', v: 0.5, i: 3, st: 0, kids: [{name: 'S1', v: 0.25, i: 4, st: 0}, {name: 'L6', v: 0.25, i: 5, st: 5}]}
];
test('one colour per top-level branch, in the order of the data, shared by its descendants', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const brush = idx => chart._getSegmentBrush({getDptByIdx: () => null, getPtBrush: () => 'series'}, idx).accent;
	assert.equal(brush(3), 0, 'B1 is the first branch of the data: accent 1');
	assert.equal(brush(4), 0);
	assert.equal(brush(5), 0);
	assert.equal(brush(0), 1, 'B2 accent 2');
	assert.equal(brush(1), 1);
	assert.equal(brush(2), 2, 'B3 accent 3');
});
test('an explicit data point fill of the file is not overridden', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const explicit = chart._getSegmentBrush({getDptByIdx: i => i === 4 ? {spPr: {Fill: {}}} : null, getPtBrush: () => 'from the file'}, 4);
	assert.equal(explicit, 'from the file');
});
// a label whose text is laid out on lines of at most eoMaxWidth: natural size w x h per line, text 'words' of width w/words
function makeLabel(idx, natural, lineH, words) {
	const lbl = {idx, recalcInfo: {}, recalculateInternal() {
		const wordW = natural / words;
		const maxW = this.eoMaxWidth || Infinity;
		const perLine = Math.min(words, Math.max(1, Math.floor(maxW / wordW)));
		const lines = Math.ceil(words / perLine);
		this.extX = Math.min(natural, perLine * wordW);
		this.extY = lines * lineH;
	}};
	lbl.recalculateInternal();
	return lbl;
}
test('a label is centred in its segment (mid angle, mid radius)', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const seg = chart.segments[3];                                   // B1: the first ring, 0.5 of the circle
	const lbl = makeLabel(3, 4, 2, 1);
	const pos = chart._calculateDLbl(lbl);
	const mid = (seg.innerRadius + seg.outerRadius) / 2, ang = seg.startAngle + seg.angle / 2;
	assert.ok(Math.abs(pos.x + 2 - (50 + mid * Math.cos(ang))) < 1e-9);
	assert.ok(Math.abs(pos.y + 1 - (50 - mid * Math.sin(ang))) < 1e-9);
	assert.equal(lbl.eoHidden, false);
	assert.notDeepEqual(plain(pos), {x: 0, y: 0});
});
test('the label reads along the radius and is never upside down', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	for (const idx of [0, 1, 2, 3, 4, 5]) {
		const seg = chart.segments[idx];
		const lbl = makeLabel(idx, 4, 2, 1);
		chart._calculateDLbl(lbl);
		const mid = seg.startAngle + seg.angle / 2;
		// direction of the text on the page (y down) is (cos r, sin r) for the clockwise rotation r
		const dx = Math.cos(lbl.eoRotation), dy = Math.sin(lbl.eoRotation);
		assert.ok(dx > -1e-9, 'text runs left to right: ' + idx);
		assert.ok(Math.abs(Math.abs(dx * Math.cos(mid) - dy * Math.sin(mid)) - 1) < 1e-9, 'along the radius: ' + idx);
	}
});
test('a label longer than the ring is wide wraps on several lines inside the segment', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const seg = chart.segments[3];                                   // B1: 0.5 of the circle, ring 1
	const ring = seg.outerRadius - seg.innerRadius;
	const lbl = makeLabel(3, ring * 1.6, 3, 4);                      // four words, 1.6 times the ring on one line
	chart._calculateDLbl(lbl);
	assert.equal(lbl.eoMaxWidth, ring, 'wrapped at the width of the ring');
	assert.ok(lbl.extY > 3, 'more than one line');
	assert.equal(lbl.eoHidden, false);
});
// a label with text of one run: the stand-in engine breaks at spaces, and inside a word when the line is narrower than
// the word; the font size (a run property) scales the width of a character, the insets of the label stay as they are
function makeTextLabel(idx, text, charW, lineH, inset) {
	inset = inset || 0;
	const words = text.split(' ');
	const run = {
		fontSize: 10, Lines: [], Content: text.split('').map(c => ({IsSpace: () => c === ' '})),
		Get_FontSize() { return this.fontSize; },
		SetFontSize(v) { this.fontSize = v; },
		protected_GetLinesCount() { return this.starts.length; },
		protected_GetRangeStartPos(line) { return this.starts[line]; }
	};
	const par = {Content: [run], CheckRunContent(fn) { fn(run); }};
	const body = {content: {Content: [par]}};
	const lbl = {idx, txBody: body, tx: {rich: body}, recalcInfo: {}, recalculateInternal() {
		const w = charW * run.fontSize / 10;
		const maxW = (this.eoMaxWidth || Infinity) - inset;
		const perLine = Math.max(1, Math.floor(maxW / w + 1e-9));    // characters per line
		const starts = [0];
		let idx = 0, cur = 0, widest = 0;
		for (const word of words) {
			let n = word.length, start = idx;
			idx += n + 1;
			if (cur > 0 && cur + 1 + n <= perLine) { cur += 1 + n; continue; }
			if (cur > 0) { starts.push(start); widest = Math.max(widest, cur); }
			while (n > perLine) { starts.push(start + perLine); widest = Math.max(widest, perLine); n -= perLine; start += perLine; }
			cur = n;
		}
		widest = Math.max(widest, cur);
		run.starts = starts;
		this.extX = widest * w + inset;
		this.extY = starts.length * lineH;
	}};
	lbl.recalculateInternal();
	return lbl;
}
test('a label wider than its ring is made smaller to stay on one line, like Excel', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const seg = chart.segments[3];
	const ring = seg.outerRadius - seg.innerRadius;
	const lbl = makeTextLabel(3, 'Stem 222', ring / 5.5, 0.5, 0);   // 8 characters, 1.45 times the ring on one line
	chart._calculateDLbl(lbl);
	assert.ok(lbl.eoFontScale < 1 && lbl.eoFontScale >= 0.65, 'smaller font');
		assert.equal(lbl.extY, 0.5, 'one line');
	assert.ok(!lbl.eoMaxWidth, 'not wrapped');
	// laid out again, the label starts from the original font again
	chart._calculateDLbl(lbl);
	assert.equal(lbl.extY, 0.5);
});
test('a wrapped label breaks at spaces only and never inside a word', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const seg = chart.segments[3];
	const ring = seg.outerRadius - seg.innerRadius;
	// 4.2 characters of the ring: even at 65% "Stem 22" (7 characters) does not fit on one line
	const charW = ring / 4.2 * 1.9;
	const lbl = makeTextLabel(3, 'Stem 22', charW, 0.5, 0);
	chart._calculateDLbl(lbl);
	const run = lbl.txBody.content.Content[0].Content[0];
	assert.equal(run.protected_GetLinesCount(), 2, 'Stem / 22');
	assert.equal(run.protected_GetRangeStartPos(1), 5, 'the second line starts after the space');
	const long = makeTextLabel(3, 'Longword 2', charW, 0.5, 0);
	chart._calculateDLbl(long);
	const lrun = long.txBody.content.Content[0].Content[0];
	assert.equal(lrun.protected_GetLinesCount(), 2, 'Longword / 2: the word stays whole');
	assert.equal(lrun.protected_GetRangeStartPos(1), 9);
});

test('a label that fits on one line is not wrapped', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const lbl = makeLabel(3, 4, 2, 2);
	chart._calculateDLbl(lbl);
	assert.ok(!lbl.eoMaxWidth);
	assert.equal(lbl.extY, 2);
});
test('a label that does not fit its segment is hidden', () => {
	const {chart} = makeChart(treeOf(SPEC), 2);
	chart.recalculate();
	const big = makeLabel(1, 60, 40, 1);
	chart._calculateDLbl(big);
	assert.equal(big.eoHidden, true);
	const tooTall = makeLabel(1, 4, 400, 1);                           // wrapped or not, three lines do not fit
	chart._calculateDLbl(tooTall);
	assert.equal(tooTall.eoHidden, true);
	const unknown = {idx: 999, extX: 1, extY: 1};
	assert.deepEqual(plain(chart._calculateDLbl(unknown)), {x: 0, y: 0});
	assert.equal(unknown.eoHidden, true);
});

test('a rotated label turns around the centre of its box, a wrapped one asks for its own width', () => {
	const ops = [];
	const mt = {
		TranslateAppend: (m, x, y) => m.ops.push(['t', x, y]), RotateRadAppend: (m, a) => m.ops.push(['r', a]),
		Invert: m => m
	};
	const mk = ops => ({ops, Reset() { this.ops = []; }, CreateDublicate() { return mk(this.ops.slice()); }});
	const matrix = () => mk([]);
	const ctx = vm.createContext({global_MatrixTransformer: mt, AscFormat: {isRealNumber: v => typeof v === 'number' && isFinite(v)}});
	const src = slice(format, 'CDLbl.prototype.setPositionRotated = function', 'CDLbl.prototype.setPosition2');
	vm.runInContext('var CDLbl = function () {};' + src + ';this.f = CDLbl.prototype.setPositionRotated;', ctx);
	const lbl = {extX: 10, extY: 4, localTransform: matrix(), ownTransformText: matrix(), setPosition(x, y) { ops.push(['pos', x, y]); }};
	ctx.f.call(lbl, 20, 30, 0.5);
	assert.deepEqual(plain(lbl.localTransform.ops), [['t', -5, -2], ['r', 0.5], ['t', 25, 32]]);
	assert.deepEqual(plain(lbl.transformText.ops), [['t', -5, -2], ['r', 0.5], ['t', 25, 32]]);
	ops.length = 0;
	ctx.f.call(lbl, 1, 2, 0);
	assert.deepEqual(plain(ops), [['pos', 1, 2]], 'no rotation: the plain position');
	const fn = slice(format, 'CDLbl.prototype.getMaxWidth = function', '\n    CDLbl.prototype.');
	const c2 = vm.createContext({AscFormat: {isRealNumber: v => typeof v === 'number' && isFinite(v)}});
	vm.runInContext('var CDLbl = function () {};' + fn + ';this.f = CDLbl.prototype.getMaxWidth;', c2);
	assert.equal(c2.f.call({eoMaxWidth: 12.5}, {}), 12.5);
});

// ------------------------------------------------------------------------------ #79: dragging a label of an extended chart
test('checkDlbl of a label whose series is a chartEx series does not throw and gives no label', () => {
	const src = slice(format, 'CDLbl.prototype.checkDlbl = function() {', 'CDLbl.prototype.checkDocContent');
	const ctx = vm.createContext({AscDFH: {historyitem_type_TrendLine: 1}, AscFormat: {}});
	vm.runInContext('var CDLbl = function () {};' + src + ';this.checkDlbl = CDLbl.prototype.checkDlbl;', ctx);
	const chartExSeries = {parent: {}};                              // no setDLbls, no dLbls: CSeries of ChartEx.js
	const lbl = {parent: null, series: chartExSeries, pt: {idx: 0}};
	assert.equal(ctx.checkDlbl.call(lbl), null);
	// a classic series still gets its dLbls
	let set = null;
	const classic = {parent: {}, setDLbls(v) { set = v; this.dLbls = v; }};
	ctx.AscFormat.CDLbls = function () { this.findDLblByIdx = () => ({}); };
	const res = ctx.checkDlbl.call({parent: null, series: classic, pt: {idx: 0}});
	assert.ok(set, 'setDLbls called for a classic series');
	assert.ok(res);
});
test('the end of a label drag in an extended chart changes nothing and does not throw', () => {
	const track0 = moves.slice(moves.indexOf('function MoveChartObjectTrack'));
	const src = slice(track0, 'this.trackEnd = function()', 'this.getBounds = function ()');
	const calls = [];
	class CDLbl {}
	const ctx = vm.createContext({AscFormat: {CDLbl}, AscCommon: {History: {Create_NewPoint: () => calls.push('point')}}, AscDFH: {}});
	const trackEnd = vm.runInContext('(function () { var o = {}; (function () {' + src + '}).call(o); return o.trackEnd; })()', ctx);
	const track = {bIsTracked: true, originalObject: new CDLbl(), chartSpace: {isChartEx: () => true}};
	assert.doesNotThrow(() => trackEnd.call(track));
	assert.deepEqual(calls, [], 'no history point for a chartEx label');
});
