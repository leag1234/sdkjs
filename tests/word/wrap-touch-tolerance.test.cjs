// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// #61: a line whose bottom touches the top of a wrapped picture (differing only by the accumulated twip rounding of
// the line heights) is not pushed under the picture.
// Run: node --test tests/word/wrap-touch-tolerance.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const src = fs.readFileSync(path.join(root, 'word/Editor/GraphicObjects/WrapManager.js'), 'utf8');

function load() {
	const ctx = vm.createContext({console});
	const consts = 'var WRAPPING_TYPE_TOP_AND_BOTTOM = 4, WRAPPING_TYPE_NONE = 0, WRAPPING_TYPE_SQUARE = 1, WRAPPING_TYPE_THROUGH = 2, WRAPPING_TYPE_TIGHT = 3;'
		+ 'var WRAP_TEXT_SIDE_BOTH_SIDES = 0, WRAP_TEXT_SIDE_LARGEST = 1, WRAP_TEXT_SIDE_LEFT = 2, WRAP_TEXT_SIDE_RIGHT = 3;';
	const tolerance = /var WRAP_TOUCH_TOLERANCE = [0-9.]+;/.exec(src);
	assert(tolerance, 'WRAP_TOUCH_TOLERANCE is defined');
	const fnStart = src.indexOf('getArrayWrapIntervals: function(x0, y0, x1, y1, LeftField, RightField, ret, bMathWrap)');
	assert(fnStart > 0);
	const fnEnd = src.indexOf('\n    },', fnStart);
	assert(fnEnd > fnStart);
	const body = src.slice(fnStart + 'getArrayWrapIntervals: '.length, fnEnd + 6);
	vm.runInContext(consts + tolerance[0] + ';var fn = ' + body + ';', ctx);
	return ctx.fn;
}

function polygon(top, bottom) {
	return {top, bottom, left: 25, right: 200, wrapSide: 0, wordGraphicObject: {useWrap: true, wrappingType: 4}};
}

test('a line touching the top of a top-and-bottom picture is not wrapped', () => {
	const fn = load();
	// the line bottom is 0.000138 mm below the picture top: the picture after a four-line paragraph
	assert.deepEqual(fn.call(polygon(75.97055664, 133.86), 25, 71.50805555, 200, 75.97069444, 25, 200, [], false), []);
});

test('a line really overlapping the picture is still wrapped', () => {
	const fn = load();
	const ret = fn.call(polygon(75.97, 133.86), 25, 71.5, 200, 76.1, 25, 200, [], false);
	assert.equal(ret.length, 1);
	assert.equal(ret[0].Y1, 133.86);
});

test('a line starting at the bottom of the picture is not wrapped', () => {
	const fn = load();
	assert.deepEqual(fn.call(polygon(75.97, 133.86), 25, 133.8599, 200, 138, 25, 200, [], false), []);
});
