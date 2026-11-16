// SPDX-License-Identifier: AGPL-3.0-only
"use strict";
// An xlsx drawing with a graphic object followed by its fallback (extended charts: mc:AlternateContent) must leave the
// stream at the end of the item. The readers of the following items go by the declared lengths, so a stream left behind
// misread the next sheets of the workbook and they were lost (a data sheet used by charts of an earlier sheet).
// Run: node --test tests/cell/drawing-item-end.test.cjs
const {test} = require("node:test"), assert = require("node:assert/strict"), fs = require("node:fs"),
	path = require("node:path"), vm = require("node:vm");
const source = fs.readFileSync(path.resolve(__dirname, "../../cell/model/Serialize.js"), "utf8");
const start = source.indexOf("this.ReadDrawing = function(type, length, oDrawing, oFlags)");
const end = source.indexOf("this.ReadPptxDrawing = function", start);
assert(start > 0 && end > start);
const body = source.slice(start, end).replace(/^this\.ReadDrawing = /, "").trim().replace(/;$/, "");

const DRAWING = {Type: 1, pptxDrawing: 7, pptxDrawingAlternative: 8, ClientData: 9};
function run({supported, consumed, itemLength}) {
	const ctx = vm.createContext({c_oSerConstants: {ReadOk: 0, ReadUnknown: 1}, c_oSer_DrawingType: DRAWING, AscFormat: {}});
	const read = vm.runInContext("(" + body + ")", ctx);
	const stream = {cur: 1000, GetUChar: () => DRAWING.pptxDrawingAlternative, GetULongLE: () => 0, Seek2(p) { this.cur = p; }};
	const self = {
		stream, bcr: {stream, Read1: () => 0},
		ReadPptxDrawing() { stream.cur += consumed; return {isSupported: () => supported}; }
	};
	const drawing = {};
	const res = read.call(self, DRAWING.pptxDrawing, itemLength, drawing, {});
	return {res, cur: stream.cur, drawing};
}

test("a supported graphic object leaves the stream at the end of its item (fallback skipped)", () => {
	const r = run({supported: true, consumed: 8043, itemLength: 8709});
	assert.equal(r.res, 0);
	assert.equal(r.cur, 1000 + 8709);
	assert(r.drawing.graphicObject);
});

test("an item read to its end is not moved", () => {
	assert.equal(run({supported: true, consumed: 500, itemLength: 500}).cur, 1500);
});

test("a stream that is already past the item end is not moved back", () => {
	assert.equal(run({supported: true, consumed: 600, itemLength: 500}).cur, 1600);
});
