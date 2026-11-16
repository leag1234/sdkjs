// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Chart ribbon, Quick Layout: asc_ChartSettings.runAsOneStep makes the setDisplay* calls of a layout one undo step.
// Without it every setDisplay* call created a history point of its own, so Word and PowerPoint needed one Ctrl+Z per
// changed element.
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const src = fs.readFileSync(path.join(root, 'common/apiCommon.js'), 'utf8');
const start = src.indexOf('\tasc_ChartSettings.prototype.setDisplayAxes = function');
const end = src.indexOf('\tasc_ChartSettings.prototype.setFUpdateGeneralChart');
assert(start > 0 && end > start);

function engine() {
	const history = {points: 0, transactions: 0, open: 0,
		Create_NewPoint() { this.points++; }, StartTransaction() { this.open++; this.transactions++; }, EndTransaction() { this.open--; }};
	const ctx = {
		Asc: {editor: {getLogicDocument: () => ({Recalculate() {}})}, c_oAscGridLinesSettings: {none: 0, major: 1, minor: 2, majorMinor: 3},
			c_oAscChartLegendShowSettings: {none: 0}},
		AscFormat: {CLineChart: function () {}, CStockChart: function () {}},
		AscCommon: {History: history}, AscDFH: {historyitem_type_ChartSpace: 1}
	};
	ctx.window = ctx;
	vm.createContext(ctx);
	vm.runInContext('function asc_ChartSettings() {}\n' + src.slice(start, end), ctx);
	const S = vm.runInContext('asc_ChartSettings', ctx);
	const axis = {setGridlinesSetting() {}};
	const chartSpace = {
		showChartTitle() {}, showDataLabels() {}, showAxisTitlesByTypes() {}, showAxesByTypes() {}, showTrendlines() {},
		getSelectedSeries: () => null, getAllSeries: () => [], chart: {setLegend() {}, createLegend() {}},
		getPlotArea: () => ({catAx: axis, valAx: axis})
	};
	const p = new S();
	p.chartSpace = chartSpace;
	p.updateChart = () => {};
	return {p, history, S};
}

const layout = p => {
	p.setDisplayChartTitle(true, false);
	p.setDisplayLegend(true, 3);
	p.setDisplayDataLabels(true, 7);
	p.setDisplayGridlines(true, false, false, false);
	p.setDisplayAxisTitles(true, false, true, false, false);
};

test('each setDisplay* call is an undo step of its own, as the right panel and the element menu need', () => {
	const {p, history} = engine();
	layout(p);
	assert.equal(history.points, 5);
});

test('the setDisplay* calls of runAsOneStep share one history point', () => {
	const {p, history} = engine();
	p.runAsOneStep(function () { layout(this); });
	assert.equal(history.points, 1);
	assert.equal(history.transactions, 1);
	assert.equal(history.open, 0, 'the transaction is closed');
});

test('runAsOneStep calls nest into the outermost point, and a failing callback leaves the state clean', () => {
	const {p, history} = engine();
	p.runAsOneStep(function () {
		this.runAsOneStep(function () { layout(this); });
		this.setDisplayChartTitle(false, false);
	});
	assert.equal(history.points, 1);
	assert.equal(history.transactions, 1);
	assert.throws(() => p.runAsOneStep(function () { throw new Error('boom'); }), /boom/);
	assert.equal(history.open, 0);
	layout(p);
	assert.equal(history.points, 1 + 1 + 5, 'after a failure the calls are separate steps again');
});

test('runAsOneStep is exported for the editors', () => {
	assert(src.includes('prot["runAsOneStep"] = prot.runAsOneStep;'));
});
