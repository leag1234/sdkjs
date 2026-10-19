/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');
const rules = {Left: 9, Center: 10, Right: 14};
const context = vm.createContext({window: {AscWord: {}, AscCommon: {c_oAscLineDrawingRule: rules}}});
vm.runInContext(fs.readFileSync(path.join(__dirname, '../../word/Editor/Table/TableDirection.js'), 'utf8'), context);
const Direction = context.window.AscWord.TableDirection;
function close(a, b) { assert(Math.abs(a - b) < 1e-9, a + ' != ' + b); }
test('first logical 90pt cell moves to the right without changing logical index or margins', () => {
	const info = {StartGridCol: 0, X_grid_start: 0, X_grid_end: 31.75, X_cell_start: 0, X_cell_end: 31.75, X_content_start: 2, X_content_end: 28};
	const result = Direction.reflectMetrics(info, 127);
	assert.equal(result.StartGridCol, 0);
	close(result.X_cell_start, 95.25); close(result.X_cell_end, 127);
	close(result.X_content_start, 99); close(result.X_content_end, 125);
	assert.equal(info.X_cell_start, 0);
});
test('reflection preserves a merged span and inverts without changing source metadata', () => {
	const info = {StartGridCol: 1, span: 2, X_grid_start: 10, X_grid_end: 70, X_cell_start: 12, X_cell_end: 68, X_content_start: 14, X_content_end: 63};
	const result = Direction.reflectMetrics(info, 100);
	assert.equal(result.span, 2); close(result.X_grid_start, 30); close(result.X_grid_end, 90);
	const restored = Direction.reflectMetrics(result, 100);
	for (const key of Object.keys(info)) close(restored[key], info[key]);
});
test('table anchor and leading indent mirror inside the text margins', () => {
	assert.equal(Direction.reflectTablePosition(72, 72, 540, 360), 180);
	assert.equal(Direction.reflectTablePosition(180, 72, 540, 360), 72);
	assert.equal(Direction.reflectTablePosition(126, 72, 540, 360), 126);
	assert.equal(Direction.reflectTablePosition(108, 72, 540, 360), 144);
});
test('physical border view exchanges sides without mutating logical properties', () => {
	const borders = {Left: {color: 'magenta'}, Right: {color: 'cyan'}, Top: {color: 'black'}};
	assert.equal(Direction.physicalBorders(borders, false), borders);
	const result = Direction.physicalBorders(borders, true);
	assert.equal(result.Left, borders.Right); assert.equal(result.Right, borders.Left);
	assert.equal(result.Top, borders.Top); assert.equal(borders.Left.color, 'magenta');
});
function graphics() {
	const result = {calls: [], drawVerLine(...args) { this.calls.push(['vertical', ...args]); },
		drawHorLineExt(...args) { this.calls.push(['horizontal', ...args]); },
		DrawEmptyTableLine(...args) { this.calls.push(['empty', ...args]); },
		p_color(...args) { assert.equal(this, result); this.calls.push(['color', ...args]); }};
	return result;
}
test('centered table sides mirror at x486/396 and vertical drawing rules swap', () => {
	const original = graphics(); const proxy = Direction.createBorderGraphics(original, 126, 360);
	proxy.drawVerLine(rules.Left, 126, 10, 30, 3);
	proxy.drawVerLine(rules.Right, 216, 10, 30, 2);
	assert.deepEqual(original.calls, [['vertical', rules.Right, 486, 10, 30, 3], ['vertical', rules.Left, 396, 10, 30, 2]]);
});
test('horizontal signed extensions swap and negate; empty line points reflect', () => {
	const original = graphics(); const proxy = Direction.createBorderGraphics(original, 126, 360);
	proxy.drawHorLineExt(rules.Center, 10, 126, 216, 2, 3, -2);
	proxy.DrawEmptyTableLine(126, 10, 216, 10);
	assert.deepEqual(original.calls, [['horizontal', rules.Center, 10, 396, 486, 2, 2, -3], ['empty', 486, 10, 396, 10]]);
});
test('passthrough state belongs to actual graphics; optional methods remain absent', () => {
	const original = graphics(); const proxy = Direction.createBorderGraphics(original, 72, 360);
	proxy.p_color(1, 2, 3, 255);
	assert.deepEqual(original.calls, [['color', 1, 2, 3, 255]]);
	assert.equal(proxy.SetBorder, undefined);
});
test('parent routing uses each table direction while retaining multi-element document-order selection', () => {
	for (const rtl of [false, true])
	for (const physicalRight of [false, true])
	{
		const document = {_tableCursorMove: {physicalRight}};
		const parent = {LogicDocument: document, Content: [{IsTable: () => true, IsRTL: () => rtl}],
			CurPos: {ContentPos: 0}, Selection: {Use: false, StartPos: 0, EndPos: 0}};
		assert.equal(Direction.parentMovesRight(parent, false), physicalRight !== rtl);
		assert.equal(Direction.parentMovesRight(parent, true), physicalRight !== rtl);
		parent.Selection = {Use: true, StartPos: 1, EndPos: 0};
		assert.equal(Direction.parentMovesRight(parent, false), false);
		assert.equal(Direction.parentMovesRight(parent, true), true);
		document._tableCursorMove = undefined;
		assert.equal(Direction.parentMovesRight(parent, false), false);
	}
});
test('physical ruler view mirrors asymmetric widths, margins and current cell and is reversible', () => {
	const logical = {X: 60, Cols: [30, 40, 50], Margins: [{Left: 1, Right: 2}, {Left: 3, Right: 4}, {Left: 5, Right: 6}], CurCol: 0};
	const before = JSON.stringify(logical);
	Direction.reflectRulerMarkup(logical, 240);
	assert.equal(logical.X, 60);
	assert.equal(JSON.stringify(logical.Cols), '[50,40,30]');
	assert.equal(logical.CurCol, 2);
	assert.equal(JSON.stringify(logical.Margins[0]), '{"Left":6,"Right":5}');
	Direction.reflectRulerMarkup(logical, 240);
	assert.equal(JSON.stringify(logical), before);
});
test('RTL ruler internal drag follows its physical boundary while retaining the leading right edge', () => {
	const markup = {X: 60, Cols: [50, 40, 30], Margins: [{Left: 1, Right: 2}, {Left: 3, Right: 4}, {Left: 5, Right: 6}], CurCol: 1};
	assert.equal(Direction.moveRtlRulerBoundary(markup, 1, 116, 210), 116);
	assert.equal(markup.X + markup.Cols.reduce((a, b) => a + b), 180);
	assert.equal(markup.X + markup.Cols[0], 116);
	assert.equal(JSON.stringify(markup.Cols), '[50,34,30]');
	Direction.reflectRulerMarkup(markup, 240);
	assert.equal(markup.X, 60);
	assert.equal(JSON.stringify(markup.Cols), '[30,34,50]');
});
test('RTL leading ruler marker moves the table and width clamps protect physical cell margins', () => {
	const markup = {X: 60, Cols: [50, 40, 30], Margins: [{Left: 1, Right: 2}, {Left: 3, Right: 4}, {Left: 5, Right: 6}], CurCol: 1};
	Direction.moveRtlRulerBoundary(markup, 3, 186, 210);
	assert.equal(markup.X, 66);
	assert.equal(JSON.stringify(markup.Cols), '[50,40,30]');
	Direction.moveRtlRulerBoundary(markup, 0, 200, 210);
	assert.equal(markup.Cols[0], 3);
	assert.equal(markup.X + markup.Cols[0], 116);
});
test('RTL outline handle is southwest and its pointer delta expands toward physical left', () => {
	const bounds = {X: 60, Y: 20, W: 120, H: 40};
	assert.equal(JSON.stringify(Direction.resizeHandleBounds(bounds, true, 2)), '{"Left":58,"Right":60,"Top":60,"Bottom":62}');
	assert.equal(JSON.stringify(Direction.resizeHandleBounds(bounds, false, 2)), '{"Left":180,"Right":182,"Top":60,"Bottom":62}');
	const start = {X: 59, Y: 61};
	assert.equal(JSON.stringify(Direction.resizePointerDelta(start, {X: 49, Y: 66}, true)), '{"W":10,"H":5}');
	assert.equal(JSON.stringify(Direction.resizePointerDelta(start, {X: 69, Y: 56}, true)), '{"W":-10,"H":-5}');
});
