// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Extended charts: regular titles, white funnel labels, Increase/Decrease/Total legend of a waterfall (like Office).
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const chartSpace = read('common/Drawings/Format/ChartSpace.js'), format = read('common/Drawings/Format/ChartFormat.js');

test('chartEx titles are not bold', () => {
	assert(format.includes('style.TextPr.Bold = !(this.isChartEx && this.isChartEx());'));
});

test('funnel labels are white unless the file sets a text colour', () => {
	assert(/chartType === AscFormat\.SERIES_LAYOUT_FUNNEL && !\(seria\.dataLabels\.txPr\)\) \{\s*this\.setSunburstLabelColor/.test(chartSpace));
});

test('waterfall legend lists Increase, Decrease, Total', () => {
	const start = chartSpace.indexOf('CChartSpace.prototype.getWaterfallLegendEntries');
	const end = chartSpace.indexOf('CChartSpace.prototype.getHierarchyLegendEntries');
	assert(start > 0 && end > start);
	const fn = chartSpace.slice(start, end).replace(/^[^=]*=\s*/, '').replace(/\/\*\*[\s\S]*$/, '');
	const run = layoutId => {
		const ctx = vm.createContext({AscFormat: {SERIES_LAYOUT_WATERFALL: 7}, AscCommon: {}});
		const f = vm.runInContext('(' + fn.trim().replace(/;$/, '') + ')', ctx);
		return JSON.parse(JSON.stringify(f.call({isChartEx: () => true, chart: {plotArea: {plotAreaRegion: {series: [{layoutId}]}}}})));
	};
	assert.deepEqual(run(7).map(e => e.name), ['Increase', 'Decrease', 'Total']);
	assert.equal(run(1), null);
});

test('waterfall bars: increase accent 1, decrease accent 2, total accent 3, data point fill wins', () => {
	const start = chartSpace.indexOf('CChartSpace.prototype.getWaterfallBrush');
	const end = chartSpace.indexOf('/** The fill of the top-level branch');
	assert(start > 0 && end > start);
	const fn = chartSpace.slice(start, end).replace(/^[^=]*=\s*/, '').replace(/\/\*\*[\s\S]*$/, '').trim().replace(/;$/, '');
	const f = vm.runInContext('(' + fn + ')', vm.createContext({}));
	const accent = (series, pt) => f.call({getBranchBrush: n => n}, series, pt);
	const series = {layoutPr: {subtotals: {idx: [5, 10]}}};
	assert.equal(accent(series, {idx: 0, val: 22}), 0);
	assert.equal(accent(series, {idx: 1, val: -12}), 1);
	assert.equal(accent(series, {idx: 5, val: 88}), 2);
	assert.equal(accent(series, {idx: 10, val: -3}), 2);
	assert.equal(accent({}, {idx: 3, val: 5}), 0);
	// the colours are set before the data point fills, and not when the series has its own fill
	const call = chartSpace.indexOf('oChartSpace.getWaterfallBrush(oSeries, oPt)');
	assert(call > 0 && call < chartSpace.indexOf('for (let nDpt = 0; nDpt < oSeries.dPt.length; ++nDpt) {\n\t\t\t\t\t\t\t\t\tlet oDpt', call));
	assert(/layoutId === AscFormat\.SERIES_LAYOUT_WATERFALL && !\(oSeries\.spPr && oSeries\.spPr\.Fill && oSeries\.spPr\.Fill\.fill\)/.test(chartSpace));
});
