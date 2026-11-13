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
