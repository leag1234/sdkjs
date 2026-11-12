// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Absolute position tab (w:ptab): model class CRunPTab, history binary, the DOCX binary items of Serialize2.js
// and the tab stop computed for it.
// Run: node --test tests/word/ptab.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');

function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}

// ---- model ----
const win = {};
const Asc = {c_oAscTabType: {Bar: 0, Clear: 1, Decimal: 2, Num: 3, Right: 4, Left: 5, Center: 6},
	c_oAscTabLeader: {None: 0, Dot: 1, Hyphen: 2, Underscore: 3, MiddleDot: 4, Heavy: 5}};
win.Asc = Asc;
win.AscWord = {CRunElementBase: function () {}};
win.AscWord.CRunElementBase.prototype.Copy = function () { return new this.constructor(); };
const ctx = vm.createContext({window: win, AscWord: win.AscWord, Asc, para_Tab: 0x15, para_PTab: 0x4f, console,
	g_oTextMeasurer: {SetTextPr() {}, SetFontSlot() {}, MeasureCode: () => ({Width: 1})}, editor: null});
vm.runInContext(read('word/Editor/Paragraph/RunContent/Tab.js'), ctx, {filename: 'Tab.js'});
const {CRunTab, CRunPTab} = win.AscWord;

test('a plain tab is not an absolute position tab, a ptab is a tab', () => {
	assert.equal(new CRunTab().IsPTab(), false);
	const p = new CRunPTab();
	assert.equal(p.IsPTab(), true);
	assert.equal(p.IsTab(), true);
	assert.equal(p.Type, 0x15);
	assert(p instanceof CRunTab);
});

test('values are validated, unknown values are ignored', () => {
	const p = new CRunPTab();
	p.SetAlignment('right'); p.SetRelativeTo('indent'); p.SetLeader('dot');
	assert.deepEqual([p.GetAlignment(), p.GetRelativeTo(), p.GetLeader()], ['right', 'indent', 'dot']);
	p.SetAlignment('bogus'); p.SetRelativeTo(''); p.SetLeader('x');
	assert.deepEqual([p.GetAlignment(), p.GetRelativeTo(), p.GetLeader()], ['right', 'indent', 'dot']);
	assert.equal(p.GetTabType(), Asc.c_oAscTabType.Right);
	assert.equal(p.GetTabLeaderType(), Asc.c_oAscTabLeader.Dot);
	p.SetAlignment('center');
	assert.equal(p.GetTabType(), Asc.c_oAscTabType.Center);
	p.SetAlignment('left');
	assert.equal(p.GetTabType(), Asc.c_oAscTabType.Left);
});

test('the recalculation calls SetLeader(type, textPr): it still sets the drawn leader', () => {
	const p = new CRunPTab();
	p.SetLeader(Asc.c_oAscTabLeader.Dot, {});
	assert.equal(p.LeaderCode, 0x2E);
	assert.equal(p.GetLeader(), 'none');
});

test('Copy keeps the attributes', () => {
	const p = new CRunPTab();
	p.SetAlignment('center'); p.SetRelativeTo('indent'); p.SetLeader('hyphen');
	const c = p.Copy();
	assert(c instanceof CRunPTab);
	assert.deepEqual([c.Alignment, c.RelativeTo, c.Leader], ['center', 'indent', 'hyphen']);
});

test('history binary round trip', () => {
	const out = [];
	const w = {WriteLong: v => out.push(['L', v]), WriteString2: v => out.push(['S', v])};
	const p = new CRunPTab();
	p.SetAlignment('right'); p.SetLeader('underscore');
	p.Write_ToBinary(w);
	assert.deepEqual(out[0], ['L', 0x4f]);
	let i = 1;
	const r = {GetString2: () => out[i++][1]};
	const q = new CRunPTab();
	q.Read_FromBinary(r);
	assert.deepEqual([q.Alignment, q.RelativeTo, q.Leader], ['right', 'margin', 'underscore']);
});

test('Types.js creates a CRunPTab for the ptab history type and a CRunTab for a tab', () => {
	const src = read('word/Editor/Paragraph/RunContent/Types.js');
	assert(src.includes('var para_PTab                      = 0x004f;'));
	assert(/case para_PTab\s*: oElement = new AscWord\.CRunPTab\(\); break;/.test(src));
	assert(/case para_Tab\s*: oElement = new AscWord\.CRunTab\(\); break;/.test(src));
});

// ---- DOCX binary items ----
const serial = read('word/Editor/Serialize2.js');
test('item numbers are the ones of the core header', () => {
	const hdr = fs.existsSync(path.join(root, '../core/OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h'))
		? read('../core/OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h') : null;
	assert(/ptab: 35,/.test(serial));
	assert(/var c_oSerPTabType = \{\s*Alignment: 0,\s*RelativeTo: 1,\s*Leader: 2\s*\};/.test(serial));
	if (hdr) {
		assert(/ptab = 35,/.test(hdr));
		assert(/Alignment = 0,\s*RelativeTo = 1,\s*Leader = 2/.test(hdr));
	}
});

test('WritePTab / ReadPTab round trip', () => {
	const code = 'var c_oSerConstants = {ReadOk: 0, ReadUnknown: 1}; var c_oSerPTabType = {Alignment: 0, RelativeTo: 1, Leader: 2};'
		+ 'var o = {};' + slice(serial, '\tthis.WritePTab = function', 'this.WriteRunContent = function').replace(/this\./g, 'o.')
		+ slice(serial, '\tthis.ReadPTab = function', '\tthis.ReadRunContent = function').replace(/this\./g, 'o.')
		+ ';this.o = o;';
	const bytes = [];
	const memory = {WriteByte: v => bytes.push(['B', v]), WriteString2: s => bytes.push(['S', s])};
	const c = vm.createContext({console});
	vm.runInContext(code, c);
	const o = c.o;
	o.memory = memory;
	o.WritePTab(Object.assign(new CRunPTab(), {Alignment: 'right', RelativeTo: 'margin', Leader: 'none'}));
	assert.deepEqual(bytes, [['B', 0], ['S', 'right'], ['B', 1], ['S', 'margin'], ['B', 2], ['S', 'none']]);
	// reading: the stream returns the value the item holds
	const p = new CRunPTab();
	let value = 'center';
	o.stream = {GetString2LE: () => value};
	assert.equal(o.ReadPTab(0, 6, p), 0);
	value = 'indent'; assert.equal(o.ReadPTab(1, 6, p), 0);
	value = 'dot'; assert.equal(o.ReadPTab(2, 3, p), 0);
	assert.deepEqual([p.Alignment, p.RelativeTo, p.Leader], ['center', 'indent', 'dot']);
	assert.equal(o.ReadPTab(9, 1, p), 1);
});

test('WritePTab is a method of the class that calls it (the run content writer), ReadPTab of the readers', () => {
	const w = serial.indexOf('\tthis.WritePTab = function');
	const r = serial.indexOf('\tthis.ReadPTab = function');
	const cls = (pos) => { const re = /\nfunction (\w+)\(/g; let m, name = null; while ((m = re.exec(serial)) && m.index < pos) name = m[1]; return name; };
	assert.equal(cls(w), cls(serial.indexOf('oThis.WritePTab(item)')));
	assert.equal(cls(r), cls(serial.indexOf('oThis.ReadPTab(t, l, oPTab)')));
});

test('the run content writer and both readers handle ptab, a plain tab keeps its item', () => {
	assert(/if \(item instanceof AscWord\.CRunPTab\) \{\s*oThis\.bs\.WriteItem\(c_oSerRunType\.ptab/.test(serial));
	assert.equal((serial.match(/c_oSerRunType\.ptab === type/g) || []).length, 2);
	assert(/else \{\s*oThis\.memory\.WriteByte\(c_oSerRunType\.tab\);/.test(serial));
});

// ---- compat flags ----
test('compat flags: unmodeled bits are written back, modeled bits come from the settings', () => {
	const code = 'var c_oSerCompat = {CompatSetting: 0, Flags1: 4, Flags2: 5, Flags3: 6};'
		+ 'var o = {};' + slice(serial, '\tthis.WriteCompat = function()', '\tthis.WriteCompatSetting = function').replace(/this\./g, 'o.')
		+ ';';
	const items = [];
	const c = vm.createContext({console, AscCommon: {document_compatibility_mode_Word15: 15}});
	vm.runInContext(code, c);
	const o = c.o;
	const settings = {BalanceSingleByteDoubleByteWidth: false, UlTrailSpace: true, DoNotExpandShiftReturn: false,
		DoNotBreakWrappedTables: false, UseFELayout: false, SplitPageBreakAndParaMark: false,
		CompatFlags: [0x80000000 | (1 << 6) | (1 << 3), (1 << 31) | (1 << 19) | (1 << 17), 1]};
	o.saveParams = {isCompatible: true};
	o.Document = {GetCompatibilityMode: () => 15, GetDocumentSettings: () => settings};
	o.memory = {WriteULong: v => items.push(v), WriteString3() {}};
	o.bs = {WriteItem: (t, f) => { if (t >= 4) { items.push('t' + t); } f(); }};
	o.WriteCompatSetting = () => {};
	o.WriteCompat();
	// flags1: bit 3 and bit 31 kept, bit 6 (modeled, false) dropped, bit 9 (modeled, true) set
	assert.deepEqual(items, ['t4', ((0x80000000 | (1 << 3) | (1 << 9)) >>> 0), 't5',
		(((1 << 31) | (1 << 19)) >>> 0), 't6', 1]);
	// without raw flags nothing but the modeled bits is written and there is no flags3 item
	items.length = 0;
	settings.CompatFlags = null;
	o.WriteCompat();
	assert.deepEqual(items, ['t4', 1 << 9, 't5', 0]);
});
