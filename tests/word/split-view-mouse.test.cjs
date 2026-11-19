// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Split window (Ctrl+Alt+S): the mouse events of the second pane (asc_MouseInPage) reach the document with the
// click count of the gesture, and the selection rectangles of a page are collected for the pane to paint.
// Run: node --test tests/word/split-view-mouse.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const source = fs.readFileSync(path.resolve(__dirname, '../../word/Editor/word-views.js'), 'utf8');
const from = source.indexOf('api["asc_MouseInPage"]'), to = source.indexOf('api["asc_WatchRecalculation"]');
assert(from > 0 && to > from);

function load(doc, dd) {
	class CMouseEventHandler { constructor() { this.Button = 0; this.ClickCount = 0; this.ShiftKey = false; this.CtrlKey = false; this.IsLocked = false; } }
	const AscCommon = {CMouseEventHandler, g_mouse_button_left: 0, g_mouse_event_type_down: 1, g_mouse_event_type_move: 2, g_mouse_event_type_up: 3};
	const api = {private_GetLogicDocument: () => doc, WordControl: {m_oDrawingDocument: dd}};
	vm.runInNewContext('(function(api){' + source.slice(from, to) + '})', {AscCommon})(api);
	return api;
}
function fakeDoc(log) {
	return {
		Selection: {DragDrop: {Flag: 1, Data: {}}}, used: false,
		IsTextSelectionUse() { return this.used; },
		IsSelectionEmpty() { return !this.used; },
		RemoveSelection() { log.push('remove'); this.used = false; },
		OnMouseDown(e, x, y, p) { log.push(['down', e.ClickCount, e.ShiftKey, e.CtrlKey, x, y, p, e.IsLocked]); this.Selection.DragDrop.Flag = 1; },
		OnMouseMove(e, x, y, p) { log.push(['move', e.ClickCount, x, y, p, e.IsLocked]); },
		OnMouseUp(e, x, y, p) { log.push(['up', e.ClickCount, x, y, p, e.IsLocked]); },
		UpdateSelection() { log.push('sel'); },
		UpdateInterface() { log.push('ui'); },
		DrawSelectionOnPage(page) { log.push(['draw', page]); dd.AddPageSelection(page, 1, 2, 3, 4); dd.AddPageSelection(page + 1, 9, 9, 9, 9); }
	};
}
let dd;

test('a gesture goes down, move, up with one click count, clamped to three', () => {
	const log = [], doc = fakeDoc(log), api = load(doc, {});
	api.asc_MouseInPage('down', 2, 10, 20, 2, false, false);
	api.asc_MouseInPage('move', 3, 11, 21, 2);
	api.asc_MouseInPage('up', 3, 12, 22, 2);
	api.asc_MouseInPage('down', 0, 1, 1, 7, true, true);
	assert.deepEqual(log.filter(Array.isArray), [
		['down', 2, false, false, 10, 20, 2, true], ['move', 2, 11, 21, 3, true], ['up', 2, 12, 22, 3, false],
		['down', 3, true, true, 1, 1, 0, true]]);
});

test('drag and drop of the selected text is not started from the second pane', () => {
	const log = [], doc = fakeDoc(log), api = load(doc, {});
	api.asc_MouseInPage('down', 0, 1, 1, 1);
	assert.equal(doc.Selection.DragDrop.Flag, 0);
	assert.equal(doc.Selection.DragDrop.Data, null);
});

test('a single click inside a selection starts a new selection; a shift click and a double click keep it', () => {
	const log = [], doc = fakeDoc(log), api = load(doc, {});
	doc.used = true;
	api.asc_MouseInPage('down', 0, 1, 1, 1, false);
	assert.deepEqual(log[0], 'remove');
	log.length = 0;
	doc.used = true;
	api.asc_MouseInPage('down', 0, 1, 1, 1, true);
	api.asc_MouseInPage('down', 0, 1, 1, 2, false);
	assert.ok(!log.includes('remove'));
});

test('nothing happens without a document or a page', () => {
	const api = load(null, {});
	assert.doesNotThrow(() => api.asc_MouseInPage('down', 0, 1, 1, 1));
	const log = [], api2 = load(fakeDoc(log), {});
	api2.asc_MouseInPage('down', -1, 1, 1, 1);
	assert.equal(log.length, 0);
});

test('the selection rectangles of one page are collected and the drawing hook is restored', () => {
	const log = [];
	const original = () => 'original';
	dd = {AddPageSelection: original};
	const doc = fakeDoc(log), api = load(doc, dd);
	assert.deepEqual(JSON.parse(JSON.stringify(api.asc_GetSelectionRects(0))), [], 'no selection, no rectangles');
	doc.used = true;
	assert.deepEqual(JSON.parse(JSON.stringify(api.asc_GetSelectionRects(0))), [{x: 1, y: 2, w: 3, h: 4}]);
	assert.equal(dd.AddPageSelection, original);
	doc.DrawSelectionOnPage = () => { throw new Error('boom'); };
	assert.throws(() => api.asc_GetSelectionRects(0), /boom/);
	assert.equal(dd.AddPageSelection, original);
});
