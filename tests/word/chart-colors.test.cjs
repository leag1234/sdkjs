// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Change Colors of the chart ribbon: asc_ChartSettings.setChartColors / getChartColorsAccent.
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const src = fs.readFileSync(path.join(root, 'common/apiCommon.js'), 'utf8');
const start = src.indexOf('\t// Change Colors.');
const end = src.indexOf('\tasc_ChartSettings.prototype.setDisplayAxes = function');
assert(start > 0 && end > start);

function engine() {
	const log = [];
	class Item { constructor(id) { this.color = {id, setId(v) { this.id = v; log.push('id' + v); }}; this.Mods = null; } setMods(m) { this.Mods = m; } }
	const base = [0, 1, 2, 3, 4, 5].map(i => new Item(i));
	const palette = () => ({getBaseColors: () => base.map(b => { const c = new Item(b.color.id); return c; }), createDuplicate() { return palette(); }});
	const colors = palette();
	let made;
	colors.createDuplicate = () => { made = palette(); const bases = made.getBaseColors(); made.getBaseColors = () => bases; return made; };
	const ctx = {
		Asc: {editor: {getLogicDocument: () => ({Recalculate() { log.push('recalc'); }})}},
		AscFormat: {
			isRealNumber: v => typeof v === 'number',
			g_oChartStyleCache: {checkColors: key => (key === 2638547519 ? colors : null)},
			ExecuteNoHistory: (f, t, a) => f.apply(t, a),
			CColorModifiers: class { constructor() { this.Mods = []; } addMod(m) { this.Mods.push(m); } },
			CColorMod: class { setName(n) { this.name = n; } setVal(v) { this.val = v; } }
		},
		AscCommon: {History: {Create_NewPoint() { log.push('point'); }}, g_oChartStyles: {}},
		AscDFH: {historyitem_type_ChartSpace: 1}
	};
	ctx.window = ctx;
	vm.createContext(ctx);
	vm.runInContext('function asc_ChartSettings() {}\n' + src.slice(start, end).replace(/\tasc_ChartSettings\.prototype\.setDisplayAxes[\s\S]*$/, ''), ctx);
	const S = vm.runInContext('asc_ChartSettings', ctx);
	return {S, log, ctx, made: () => made};
}

test('the monochromatic palette is built from the colorful one: six shades and tints of one accent', () => {
	const {S, log, made} = engine();
	let chosen;
	const chartSpace = {chartStyle: {}, chartColors: null, setChartColors(c) { chosen = c; chartSpace.chartColors = c; }, resetToChartStyle() { log.push('reset'); }};
	const p = new S();
	p.chartSpace = chartSpace;
	p.updateChart = () => log.push('update');
	p.setChartColors(2);
	assert.equal(chosen, made());
	const bases = made().getBaseColors();
	assert.deepEqual(bases.map(b => b.color.id), [2, 2, 2, 2, 2, 2]);
	assert.deepEqual(bases.map(b => b.Mods.Mods.map(m => m.name + m.val).join('+')), ['shade76000', 'shade88000', '', 'tint77000', 'tint60000', 'tint45000']);
	assert.deepEqual(log.filter(l => !/^id/.test(l)), ['point', 'reset', 'update', 'recalc']);
});

test('Colorful keeps the accents in turn and an unknown state reads as -1', () => {
	const {S, made} = engine();
	const chartSpace = {chartStyle: {}, chartColors: null, setChartColors(c) { chartSpace.chartColors = c; }, resetToChartStyle() {}};
	const p = new S();
	p.chartSpace = chartSpace;
	p.updateChart = () => {};
	p.setChartColors(-1);
	assert.deepEqual(made().getBaseColors().map(b => b.color.id), [0, 1, 2, 3, 4, 5]);
	assert.equal(p.getChartColorsAccent(), -1);
	p.setChartColors(4);
	assert.equal(p.getChartColorsAccent(), 4);
});

test('a chart without colours reports null, and one without a style part starts from the first style of its type', () => {
	const {S, log, ctx} = engine();
	const p = new S();
	assert.equal(p.getChartColorsAccent(), null);
	p.chartSpace = {chartColors: null};
	assert.equal(p.getChartColorsAccent(), null);
	p.chartSpace = null;
	p.setChartColors(0);
	assert.equal(log.length, 0, 'nothing happens without a chart');
	const applied = [];
	const chartSpace = {chartStyle: null, chartColors: null, getChartType: () => 7, applyChartStyleByIds(ids) { applied.push(ids); }, setChartColors() {}, resetToChartStyle() {}};
	p.chartSpace = chartSpace;
	p.updateChart = () => {};
	ctx.AscCommon.g_oChartStyles[7] = [['style-a'], ['style-b']];
	p.setChartColors(-1);
	assert.deepEqual(JSON.parse(JSON.stringify(applied)), [['style-a']]);
});

test('the methods are exported for the editors', () => {
	assert(src.includes('prot["getChartColorsAccent"] = prot.getChartColorsAccent;'));
	assert(src.includes('prot["setChartColors"] = prot.setChartColors;'));
});
