'use strict';
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const win = {AscCommon: {}};
vm.runInContext(fs.readFileSync(path.join(root, 'common/EoSvgHairline.js'), 'utf8'), vm.createContext({window: win}), {filename: 'EoSvgHairline.js'});
const scale = win.AscCommon.EoSvgEffectiveScale;

test('media1 pattern becomes a hairline', () => {
	assert.ok(Math.abs(scale(['scale(0.021739,0.021739)', 'scale(0.90664,0.90456)'], 1) - 0.0197) < 0.0005);
});

test('normal stroke is unchanged', () => {
	assert.equal(scale([], 1), 1);
});

test('matrix determinant is supported', () => {
	assert.equal(scale(['matrix(2,0,0,3,10,20)'], 1), Math.sqrt(6));
});

test('malformed SVG is returned unchanged', () => {
	const malformed = '<svg><polyline stroke="black"';
	assert.equal(win.AscCommon.EoFixSvgHairlines(malformed), malformed);
});

test('hairline width is 0.6 px, scaled by the pixel ratio', () => {
	assert.equal(win.AscCommon.EoSvgHairlinePx(), 0.6);
	win.AscCommon.AscBrowser = {retinaPixelRatio: 2};
	assert.equal(win.AscCommon.EoSvgHairlinePx(), 1.2);
	delete win.AscCommon.AscBrowser;
});

test('PDF hairline is 0.14 pt (Word) at the raster resolution', () => {
	assert.ok(Math.abs(win.AscCommon.EoSvgHairlinePxForDpi(300) - 0.5833) < 0.001);
	assert.ok(Math.abs(win.AscCommon.EoSvgHairlinePxForDpi(72) - 0.14) < 1e-9);
});

test('the PDF export applies the hairline fix to SVG pictures', () => {
	const pdf = fs.readFileSync(path.join(root, 'common/clientpdf/clientpdf.js'), 'utf8');
	assert.ok(pdf.includes('AscCommon.EoFixSvgHairlines(svgText, { minPx: AscCommon.EoSvgHairlinePxForDpi(300), destWidthPx: w })'));
});
