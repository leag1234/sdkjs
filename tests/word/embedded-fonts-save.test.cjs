// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Fonts embedded in DOCX are kept on save: the font table of the saved binary (BinaryFontTableWriter), the key and
// subsetted items read back by Binary_FontTableReader, and the hook in the main table writer.
// Run: node --test tests/word/embedded-fonts-save.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const serial = read('word/Editor/Serialize2.js');

function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}
const eq = (a, b, m) => assert.equal(JSON.stringify(a), JSON.stringify(b), m);

const win = {AscCommon: {}, Asc: {}};
const commonCtx = vm.createContext({
	window: win, AscCommon: win.AscCommon, Asc: win.Asc, console, Uint8Array, TextDecoder,
	AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}}
});
vm.runInContext(read('common/SerializeCommonWordExcel.js'), commonCtx, {filename: 'SerializeCommonWordExcel.js'});
const AC = win.AscCommon;

// a minimal CMemory: growing byte array with the calls the table writers use
function Memory() {
	this.b = [];
	this.p = 0;
	this.put = v => { this.b[this.p++] = v & 0xFF; };
	this.GetCurPosition = () => this.p;
	this.Seek = p => { this.p = p; };
	this.Skip = n => { for (let i = 0; i < n; ++i) this.put(0); };
	this.WriteByte = v => this.put(v);
	this.WriteLong = v => { this.put(v); this.put(v >> 8); this.put(v >> 16); this.put(v >> 24); };
	this.WriteString3 = s => { for (let i = 0; i < s.length; ++i) { this.put(s.charCodeAt(i)); this.put(s.charCodeAt(i) >> 8); } };
}

const code = slice(serial, 'var c_oSerFontTableTypes = {', 'var c_oSerSigTypes')
	+ slice(serial, 'function Binary_FontTableReader(', 'function Binary_SettingsTableReader(')
	+ slice(serial, 'function BinaryFontTableWriter(', 'function BinarySigTableWriter(')
	+ ';this.Types = c_oSerFontTableTypes; this.Reader = Binary_FontTableReader; this.Writer = BinaryFontTableWriter;';
const ctx = vm.createContext({
	AscCommon: AC, Asc: win.Asc, console, c_oSerConstants: AC.c_oSerConstants, c_oSerPropLenType: AC.c_oSerPropLenType,
	Binary_CommonReader: AC.Binary_CommonReader, BinaryCommonWriter: AC.BinaryCommonWriter
});
vm.runInContext(code, ctx, {filename: 'Serialize2.js'});

const KEY1 = '{3A7E6E5F-1F4B-4D2D-A1B2-0123456789AB}', KEY2 = '{00112233-4455-6677-8899-AABBCCDDEEFF}';
const FACES = [
	{name: 'Zz Test', slot: 0, path: 'fonts/Zz_Test_regular.ttf', key: KEY1, subsetted: true},
	{name: 'Zz Test', slot: 3, path: 'fonts/Zz_Test_bolditalic.ttf', key: KEY2, subsetted: false},
	{name: 'Other', slot: 1, path: 'fonts/Other_bold.otf', key: KEY2, subsetted: false}
];
function write(faces) {
	const m = new Memory();
	const w = new ctx.Writer(m, faces);
	w.Write();
	return {bytes: Uint8Array.from(m.b), writer: w};
}
function readBack(bytes) {
	const result = {fontAltNames: {}, fontFamilies: {}, embeddedFonts: []};
	const res = new ctx.Reader(result, new AC.FT_Stream2(bytes, bytes.length)).Read();
	return {res, result};
}

test('item numbers of the key and subsetted items equal core BinReaderWriterDefines.h', () => {
	const core = path.resolve(root, '../core/OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h');
	if (!fs.existsSync(core))
		return;
	const header = slice(fs.readFileSync(core, 'utf8'), 'namespace c_oSerFontTableTypes', '};}');
	for (const k of ['EmbedKeyRegular', 'EmbedKeyBold', 'EmbedKeyItalic', 'EmbedKeyBoldItalic', 'EmbedSubsetted']) {
		const m = new RegExp('\\b' + k + '\\s*=\\s*(\\d+)').exec(header);
		assert(m, 'core has ' + k);
		assert.equal(+m[1], ctx.Types[k], k);
	}
});

test('write then read keeps name, slot, path, key and subsetted of every embedded face', () => {
	const {bytes} = write(FACES);
	const {res, result} = readBack(bytes);
	assert.equal(res, AC.c_oSerConstants.ReadOk);
	eq(result.embeddedFonts, FACES);
});

test('the table has the layout core reads: table length, Font items with length, strings as UTF-16', () => {
	const {bytes} = write([FACES[0]]);
	const dv = new DataView(bytes.buffer);
	assert.equal(dv.getInt32(0, true), bytes.length - 4);
	assert.equal(bytes[4], 0);							// Font
	assert.equal(dv.getInt32(5, true), bytes.length - 9);
	assert.equal(bytes[9], 1);							// Name
	const nameLen = dv.getInt32(10, true);
	assert.equal(nameLen, 'Zz Test'.length * 2);
	let pos = 14 + nameLen;
	const types = [];
	while (pos < bytes.length) {
		types.push(bytes[pos]);
		pos += 5 + dv.getInt32(pos + 1, true);
	}
	eq(types, [4, 8, 12]);
	assert.equal(pos, bytes.length);
	// the subsetted item is one byte with the slot mask
	assert.equal(bytes[bytes.length - 1], 1);
});

test('faces without a path, key or valid slot are not written; no table at all without faces', () => {
	const w = new ctx.Writer(new Memory(), [
		{name: 'A', slot: 0, path: 'fonts/a.ttf', key: ''}, {name: 'A', slot: 1, path: '', key: KEY1},
		{name: 'A', slot: 7, path: 'fonts/a.ttf', key: KEY1}, {name: '', slot: 0, path: 'fonts/a.ttf', key: KEY1}
	]);
	assert.equal(w.HasFaces(), false);
	assert.equal(new ctx.Writer(new Memory(), []).HasFaces(), false);
});

test('a key that is not a GUID is dropped by the reader (the face is used but not saved)', () => {
	const m = new Memory();
	const bs = new AC.BinaryCommonWriter(m);
	const nStart = bs.WriteItemWithLengthStart();
	bs.WriteItem(0, () => {
		bs.WriteItem(1, () => m.WriteString3('Zz'));
		bs.WriteItem(4, () => m.WriteString3('fonts/zz.ttf'));
		bs.WriteItem(8, () => m.WriteString3('not-a-guid'));
		bs.WriteItem(5, () => m.WriteString3('fonts/zzb.ttf'));
		bs.WriteItem(9, () => m.WriteString3('3A7E6E5F1F4B4D2DA1B20123456789AB'));
	});
	bs.WriteItemWithLengthEnd(nStart);
	const {result} = readBack(Uint8Array.from(m.b));
	eq(result.embeddedFonts.map(f => f.key), ['', '3A7E6E5F1F4B4D2DA1B20123456789AB']);
});

test('the main table writer writes the font table only when the document carries embedded fonts', () => {
	const src = slice(serial, 'this.WriteMainTableContent = function', 'this.WriteMainTableEnd = function');
	assert(/this\.Document\.EmbeddedFontsInfo/.test(src));
	assert(/if \(embeddedFonts && embeddedFonts\.length > 0 && new BinaryFontTableWriter\(this\.memory, embeddedFonts\)\.HasFaces\(\)\)/.test(src));
	// a single write site: the copy/paste writer does not write the table
	assert.equal(serial.split('this.WriteTable(c_oSerTableTypes.FontTable').length - 1, 1);
});

test('the reader of a file keeps the faces on the document for the next save, and not for copy/paste', () => {
	const src = slice(serial, 'if (!this.oReadResult.bCopyPaste && AscFonts.g_fontApplication) {', 'if(null !== this.oReadResult.defaultTabStop');
	assert(/this\.Document\.EmbeddedFontsInfo = this\.oReadResult\.embeddedFonts\.length > 0 \? this\.oReadResult\.embeddedFonts\.slice\(\) : null/.test(src));
	assert(/if \(!this\.oReadResult\.bCopyPaste\) \{\s*\/\/ the embedded fonts of the file/.test(src));
});

test('a document without embedded fonts: the binary is unchanged (no FontTable table written)', () => {
	// the writer hook is the only change in the main table; with no faces its condition is false
	const src = slice(serial, 'this.WriteMainTableContent = function', 'this.WriteMainTableEnd = function');
	const hook = slice(src, 'let embeddedFonts', 'let oform');
	let written = 0;
	const run = faces => {
		written = 0;
		const doc = {EmbeddedFontsInfo: faces};
		const f = new Function('BinaryFontTableWriter', 'c_oSerTableTypes', 'ctx', 'self', 'const t = self; ' + hook.replace(/this\./g, 'self.'));
		f(ctx.Writer, {FontTable: 22}, ctx, {Document: doc, memory: new Memory(), WriteTable: () => { ++written; }});
		return written;
	};
	assert.equal(run(null), 0);
	assert.equal(run(undefined), 0);
	assert.equal(run([]), 0);
	assert.equal(run([{name: 'A', slot: 0, path: 'fonts/a.ttf', key: ''}]), 0);
	assert.equal(run(FACES), 1);
});
