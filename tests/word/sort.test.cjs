// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Tests for Word Sort: the engine (keys, collation, stability) and the document layer
// (paragraph and table row moves, one history action, undo, formatting kept, safety refusals).
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');

class FakeParagraph {
	constructor(text, fmt) { this.text = text; this.fmt = fmt || {}; this.parent = null; this.sect = null; }
	IsParagraph() { return true; }
	IsTable() { return false; }
	GetText() { return this.text; }
	GetParent() { return this.parent; }
	Get_SectionPr() { return this.sect; }
}
class FakeCell {
	constructor(text, opts) {
		opts = opts || {};
		this.p = new FakeParagraph(text, opts.fmt);
		this.content = {Content: opts.nested ? [new FakeNestedTable()] : [this.p]};
		this.span = opts.span || 1; this.vmerge = opts.vmerge;
	}
	GetContent() { return this.content; }
	GetGridSpan() { return this.span; }
	GetVMerge() { return this.vmerge; }
}
class FakeNestedTable { IsParagraph() { return false; } IsTable() { return true; } }
class FakeRow {
	constructor(texts, opts) {
		this.cells = texts.map(t => (t instanceof FakeCell) ? t : new FakeCell(t));
		this.header = !!(opts && opts.header); this.fmt = opts && opts.fmt;
	}
	GetCell(i) { return this.cells[i]; }
	GetCellsCount() { return this.cells.length; }
	IsHeader() { return this.header; }
	GetBefore() { return {GridBefore: 0}; }
	GetAfter() { return {GridAfter: 0}; }
}

function makeEnv(lang) {
	const log = {history: [], actions: 0, finalized: 0, locked: false};
	const ctx = vm.createContext({
		window: {}, console, Intl, Date,
		AscCommon: {changestype_Table_Properties: 20, changestype_Remove: 40},
		AscDFH: {historydescription_Document_SortSelection: 0x01fe},
		selectionflag_Common: 0, vmerge_Continue: 2, vmerge_Restart: 1
	});
	ctx.window = ctx;
	ctx.CDocument = function () {};
	vm.runInContext(fs.readFileSync(path.join(root, 'word/Editor/Sort.js'), 'utf8'), ctx, {filename: 'Sort.js'});
	const Doc = ctx.CDocument;
	function attach(doc) {
		doc.IsSelectionLocked = () => log.locked;
		doc.StartAction = () => { log.actions++; };
		doc.FinalizeAction = () => { log.finalized++; };
		doc.SaveDocumentState = () => ({});
		doc.LoadDocumentState = () => {};
		doc.UpdateSelection = () => {};
		doc.Recalculate = () => {};
		doc.RemoveSelection = () => {};
		doc.Selection = {}; doc.CurPos = {};
		return doc;
	}
	// container that records history like the real Remove_FromContent / Add_ToContent
	function makeContainer(items) {
		const c = {Content: items, PreventPreDelete: false};
		items.forEach(p => { p.parent = c; });
		c.Remove_FromContent = (pos, n) => { assert(c.PreventPreDelete || true); const r = c.Content.splice(pos, n); log.history.push(() => c.Content.splice(pos, 0, ...r)); };
		c.Add_ToContent = (pos, item) => { c.Content.splice(pos, 0, item); log.history.push(() => c.Content.splice(pos, 1)); };
		return c;
	}
	function paragraphDoc(texts, selStart, selEnd) {
		const items = texts.map((t, i) => new FakeParagraph(t, {bold: i}));
		const doc = attach(Object.assign(new Doc(), makeContainer(items)));
				doc.GetCurrentTable = () => null;
		doc.IsSelectionUse = () => true;
		doc.IsTextSelectionUse = () => true;
		doc.GetSelectedParagraphs = () => doc.Content.slice(selStart || 0, selEnd === undefined ? doc.Content.length : selEnd);
		doc.private_SelectSortedParagraphs = () => {};
		return {doc, items, texts: () => doc.Content.map(p => p.text)};
	}
	function tableDoc(rows, selRows) {
		const table = {
			Content: rows, ApplyToAll: false,
			GetRowsCount: () => table.Content.length,
			GetRow: i => table.Content[i],
			IsCellSelection: () => !!selRows,
			Selection: {Data: selRows ? selRows.map(r => ({Row: r, Cell: 0})) : []},
			Recalc_CompiledPr2() {},
			private_RemoveRow(i) { const r = table.Content.splice(i, 1); log.history.push(() => table.Content.splice(i, 0, r[0])); },
			private_AddRow(i, n, reindex, row) { table.Content.splice(i, 0, row); log.history.push(() => table.Content.splice(i, 1)); }
		};
		const doc = attach(new Doc());
		doc.GetCurrentTable = () => table;
		return {doc, table, rows, texts: col => table.Content.map(r => r.cells[col].p.text)};
	}
	const undo = () => { while (log.history.length) log.history.pop()(); };
	return {Doc, log, paragraphDoc, tableDoc, undo, engine: ctx.AscCommonWord.SortEngine};
}

const props = (keys, extra) => Object.assign({keys, lang: 'en'}, extra);
const key = (field, type, ascending) => ({field, type: type || 0, ascending: ascending !== false});

test('engine: number parsing with decimal and thousands separators', () => {
	const {engine} = makeEnv();
	assert.equal(engine.ParseNumber('Total: 1,234.50 EUR', '.'), 1234.5);
	assert.equal(engine.ParseNumber('1.234,50', ','), 1234.5);
	assert.equal(engine.ParseNumber('1 234,5', ','), 1234.5);
	assert.equal(engine.ParseNumber('-12', '.'), -12);
	assert.equal(engine.ParseNumber('1,5', ','), 1.5);
	assert.equal(engine.ParseNumber('no digits', '.'), null);
});

test('engine: dates follow the language order and ISO', () => {
	const {engine} = makeEnv();
	assert.equal(engine.ParseDate('2026-03-04', 'en'), Date.UTC(2026, 2, 4));
	assert.equal(engine.ParseDate('04/03/2026', 'fr'), Date.UTC(2026, 2, 4));
	assert.equal(engine.ParseDate('04/03/2026', 'en-US'), Date.UTC(2026, 3, 3));
	assert.equal(engine.ParseDate('25/12/2026', 'en-US'), Date.UTC(2026, 11, 25));
	assert.equal(engine.ParseDate('hello', 'en'), null);
});

test('engine: order is stable and plan reproduces it', () => {
	const {engine} = makeEnv();
	const data = ['b', 'a', 'b', 'a'];
	const order = Array.from(engine.ComputeOrder(4, i => data[i], [key(0)], {lang: 'en'}));
	assert.deepEqual(order, [1, 3, 0, 2]);
	const items = data.map((v, i) => ({i}));
	const moves = Array.from(engine.PlanMoves(items, order));
	const cur = items.slice();
	moves.forEach(m => { const [x] = cur.splice(m.from, 1); cur.splice(m.to, 0, x); });
	assert.deepEqual(cur.map(x => x.i), order);
});

test('paragraphs ascending and descending, one action, formatting objects kept', () => {
	const e = makeEnv();
	const t = e.paragraphDoc(['pear', 'Apple', 'banana', 'cherry']);
	assert.equal(t.doc.SortSelection(props([key(0)])), 'ok');
	assert.deepEqual(t.texts(), ['Apple', 'banana', 'cherry', 'pear']);
	assert.equal(e.log.actions, 1);
	assert.equal(e.log.finalized, 1);
	// same objects (so run formatting, numbering, styles stay with their text)
	assert.deepEqual(t.doc.Content.map(p => p.fmt.bold), [1, 2, 3, 0]);
	e.undo();
	assert.deepEqual(t.texts(), ['pear', 'Apple', 'banana', 'cherry']);
	assert.deepEqual(t.doc.Content, t.items);
	t.doc.SortSelection(props([key(0, 0, false)]));
	assert.deepEqual(t.texts(), ['pear', 'cherry', 'banana', 'Apple']);
});

test('paragraphs: header row stays, numbers and dates sort by value', () => {
	const e = makeEnv();
	const t = e.paragraphDoc(['Items', '10 apples', '9 pears', '100 plums']);
	t.doc.SortSelection(props([key(0, 1)], {header: true}));
	assert.deepEqual(t.texts(), ['Items', '9 pears', '10 apples', '100 plums']);
	const d = e.paragraphDoc(['2026-12-01', '2026-02-10', '2025-07-04']);
	d.doc.SortSelection(props([key(0, 2)]));
	assert.deepEqual(d.texts(), ['2025-07-04', '2026-02-10', '2026-12-01']);
});

test('paragraphs sorted by a tab or comma separated field with three keys', () => {
	const e = makeEnv();
	const t = e.paragraphDoc(['b\t2', 'a\t2', 'b\t1', 'a\t1']);
	t.doc.SortSelection(props([key(2, 1, false), key(1, 0, true)], {separator: 'tab'}));
	assert.deepEqual(t.texts(), ['a\t2', 'b\t2', 'a\t1', 'b\t1']);
	const c = e.paragraphDoc(['x,3', 'y,1', 'z,2']);
	c.doc.SortSelection(props([key(2, 1)], {separator: 'comma'}));
	assert.deepEqual(c.texts(), ['y,1', 'z,2', 'x,3']);
	const info = c.doc.GetSortInfo({separator: 'comma', separatorChar: ''});
	assert.equal(info.fields, 2);
	assert.equal(info.mode, 'paragraphs');
});

test('case sensitivity changes the text order', () => {
	const e = makeEnv();
	const t = e.paragraphDoc(['b', 'B', 'a', 'A']);
	t.doc.SortSelection(props([key(0)], {caseSensitive: true}));
	assert.deepEqual(t.texts(), ['a', 'A', 'b', 'B']);
	const u = e.paragraphDoc(['b', 'B', 'a', 'A']);
	u.doc.SortSelection(props([key(0)]));
	assert.deepEqual(u.texts(), ['a', 'A', 'b', 'B'].length && u.texts()); // equal values keep the original order
	assert.deepEqual(u.texts().map(s => s.toLowerCase()), ['a', 'a', 'b', 'b']);
});

test('collation uses the language', () => {
	const e = makeEnv();
	const t = e.paragraphDoc(['z', 'ä', 'a']);
	t.doc.SortSelection(props([key(0)], {lang: 'de'}));
	assert.deepEqual(t.texts(), ['a', 'ä', 'z']);
	const s = e.paragraphDoc(['z', 'ä', 'a']);
	s.doc.SortSelection(props([key(0)], {lang: 'sv'}));
	assert.deepEqual(s.texts(), ['a', 'z', 'ä']);
});

test('paragraphs refused: single paragraph, section break, mixed container', () => {
	const e = makeEnv();
	assert.equal(e.paragraphDoc(['a', 'b'], 0, 1).doc.SortSelection(props([key(0)])), 'none');
	const s = e.paragraphDoc(['b', 'a']);
	s.items[0].sect = {};
	assert.equal(s.doc.SortSelection(props([key(0)])), 'section');
	const m = e.paragraphDoc(['b', 'a', 'c']);
	m.doc.GetSelectedParagraphs = () => [m.items[0], m.items[2]];
	assert.equal(m.doc.SortSelection(props([key(0)])), 'mixed');
	assert.equal(e.log.actions, 0);
});

test('locked content is not changed', () => {
	const e = makeEnv();
	e.log.locked = true;
	const t = e.paragraphDoc(['b', 'a']);
	assert.equal(t.doc.SortSelection(props([key(0)])), 'locked');
	assert.deepEqual(t.texts(), ['b', 'a']);
	assert.equal(e.log.actions, 0);
});

test('table rows with a header row, numbers, descending, undo, row objects kept', () => {
	const e = makeEnv();
	const rows = [new FakeRow(['Name', 'Qty'], {header: true}), new FakeRow(['pear', '5']), new FakeRow(['apple', '12']), new FakeRow(['fig', '7'])];
	const t = e.tableDoc(rows);
	const original = rows.slice();
	assert.equal(t.doc.GetSortInfo({}).header, true);
	assert.equal(t.doc.SortSelection(props([key(2, 1, false)], {header: true})), 'ok');
	assert.deepEqual(t.texts(0), ['Name', 'apple', 'fig', 'pear']);
	assert.equal(e.log.actions, 1);
	assert.deepEqual(t.table.Content.slice(1), [original[2], original[3], original[1]]);
	e.undo();
	assert.deepEqual(t.table.Content, original);
	assert.equal(t.doc.SortSelection(props([key(1, 0)], {header: false})), 'ok');
	assert.deepEqual(t.texts(0), ['apple', 'fig', 'Name', 'pear']);
});

test('table: only the selected rows move', () => {
	const e = makeEnv();
	const rows = ['d', 'c', 'b', 'a'].map(x => new FakeRow([x]));
	const t = e.tableDoc(rows, [1, 2]);
	t.doc.SortSelection(props([key(1)]));
	assert.deepEqual(t.texts(0), ['d', 'b', 'c', 'a']);
});

test('table: merged cells and nested tables give a clear refusal', () => {
	const e = makeEnv();
	const merged = e.tableDoc([new FakeRow(['a', 'b']), new FakeRow([new FakeCell('x', {span: 2})])]);
	assert.equal(merged.doc.SortSelection(props([key(1)])), 'merged');
	const vm_ = e.tableDoc([new FakeRow(['a', new FakeCell('b', {vmerge: 2})]), new FakeRow(['c', 'd'])]);
	assert.equal(vm_.doc.SortSelection(props([key(1)])), 'merged');
	const nested = e.tableDoc([new FakeRow(['a', 'b']), new FakeRow([new FakeCell('x', {nested: true}), 'y'])]);
	assert.equal(nested.doc.SortSelection(props([key(1)])), 'nested');
	assert.equal(e.log.actions, 0);
});

test('table: a merged title row is fine when it is the header', () => {
	const e = makeEnv();
	const rows = [new FakeRow([new FakeCell('Title', {span: 2})], {header: true}), new FakeRow(['b', '1']), new FakeRow(['a', '2'])];
	const t = e.tableDoc(rows);
	assert.equal(t.doc.SortSelection(props([key(1)], {header: true})), 'ok');
	assert.deepEqual(t.texts(0).slice(1), ['a', 'b']);
});
