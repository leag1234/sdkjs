// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// w:autoSpaceDE / w:autoSpaceDN on paragraph properties: CParaPr model (Copy, Merge, Is_Equal, defaults,
// history binary round trip) and the DOCX binary items 50/51 of Serialize2.js (writer emits them only when
// set, reader restores them).
// Run: node --test tests/word/autospace.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');

function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = to === undefined ? source.length : source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}

// ---- common serialization code of sdkjs (BinaryCommonWriter, Binary_CommonReader, FT_Stream2) ----
const win = {AscCommon: {}, Asc: {}};
const commonCtx = vm.createContext({
	window: win, AscCommon: win.AscCommon, Asc: win.Asc, console, Uint8Array, TextDecoder,
	AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}}
});
vm.runInContext(read('common/SerializeCommonWordExcel.js'), commonCtx, {filename: 'SerializeCommonWordExcel.js'});
const AC = win.AscCommon;

function Memory() {
	this.data = [];
	this.pos = 0;
	this.size = 0;
	this.GetCurPosition = () => this.pos;
	this.Skip = n => { for (let i = 0; i < n; ++i) this.WriteByte(0); };
	this.Seek = n => { this.pos = n; };
	this.WriteByte = v => { this.data[this.pos++] = v & 0xFF; this.size = Math.max(this.size, this.pos); };
	this.WriteBool = v => this.WriteByte(v ? 1 : 0);
	this.WriteLong = v => { for (let i = 0; i < 4; ++i) this.WriteByte((v >> (8 * i)) & 0xFF); };
	this.WriteDouble = v => { const b = new ArrayBuffer(8); new DataView(b).setFloat64(0, v, true);
		new Uint8Array(b).forEach(x => this.WriteByte(x)); };
	this.WriteString2 = s => {
		this.WriteLong(s.length * 2);
		for (let i = 0; i < s.length; ++i) {
			const c = s.charCodeAt(i);
			this.WriteByte(c & 0xFF);
			this.WriteByte(c >> 8);
		}
	};
	this.bytes = () => Uint8Array.from(this.data.slice(0, this.size));
}

// ---- CParaPr from Styles.js ----
function loadParaPr() {
	const stubs = {};
	const mk = name => function () {
		this.Copy = () => new stubs[name]();
		this.Set_FromObject = () => {};
		this.Merge = () => {};
		this.Write_ToBinary = () => {};
		this.Read_FromBinary = () => {};
		this.SetLineTwips = () => {};
	};
	stubs.ParaSpacing = mk('ParaSpacing');
	stubs.CParaInd = mk('CParaInd');
	stubs.CDocumentShd = mk('CDocumentShd');
	stubs.CDocumentBorder = mk('CDocumentBorder');
	stubs.CParaTabs = mk('CParaTabs');
	const ctx = vm.createContext({
		AscWord: {ParaSpacing: stubs.ParaSpacing, CParaInd: stubs.CParaInd,
			CDocumentShd: stubs.CDocumentShd, CDocumentBorder: stubs.CDocumentBorder},
		AscCommon: AC, Asc: win.Asc,
		CParaInd: stubs.CParaInd, CDocumentShd: stubs.CDocumentShd, CDocumentBorder: stubs.CDocumentBorder,
		CParaTabs: stubs.CParaTabs,
		align_Left: 0, align_Center: 1, align_Right: 2, align_Justify: 3,
		linerule_Auto: 0,
		IsEqualStyleObjects: () => true,
		console
	});
	const src = read('word/Editor/Styles.js');
	let code = slice(src, 'function CParaPr()', 'CParaPr.prototype.Get_PresentationBullet');
	code = code.slice(0, code.lastIndexOf('/**'));
	code += ';this.CParaPr = CParaPr;';
	vm.runInContext(code, ctx, {filename: 'Styles.js'});
	return ctx.CParaPr;
}
const CParaPr = loadParaPr();

// ---- Binary_pPrWriter / Binary_pPrReader from Serialize2.js ----
const sandbox = {
	AscCommon: AC, Asc: win.Asc, AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}},
	c_oSerPropLenType: AC.c_oSerPropLenType, c_oSerConstants: AC.c_oSerConstants,
	BinaryCommonWriter: AC.BinaryCommonWriter, Binary_CommonReader: AC.Binary_CommonReader,
	align_Left: 0, align_Center: 1, align_Right: 2, align_Justify: 3,
	Binary_rPrWriter: function () { this.Write_rPr = () => {}; },
	Binary_rPrReader: function () {},
	WriteTrackRevision: () => {},
	console
};
function loadSerialize() {
	const ctx = vm.createContext(Object.assign({}, sandbox));
	const src = read('word/Editor/Serialize2.js');
	const code = slice(src, 'var c_oSerProp_pPrType = {', 'var c_oSerProp_rPrType = {')
		+ slice(src, 'function Binary_pPrWriter(', 'function Binary_rPrWriter(')
		+ slice(src, 'function Binary_pPrReader(', 'function Binary_rPrReader(')
		+ ';this.pPrType = c_oSerProp_pPrType; this.PPrWriter = Binary_pPrWriter; this.PPrReader = Binary_pPrReader;';
	vm.runInContext(code, ctx, {filename: 'Serialize2.js'});
	return ctx;
}
const ser = loadSerialize();

function writePPr(pPr) {
	const memory = new Memory();
	const w = new ser.PPrWriter(memory, null, null, {});
	w.Write_pPr(pPr, null, null, null, null);
	return memory.bytes();
}
function readPPr(bytes) {
	const pPr = new CParaPr();
	const stream = new AC.FT_Stream2(bytes, bytes.length);
	const r = new ser.PPrReader(null, null, stream);
	r.Read(bytes.length, pPr, null, null);
	return pPr;
}

test('enum values are appended at the end', () => {
	assert.equal(ser.pPrType.AutoSpaceDE, 50);
	assert.equal(ser.pPrType.AutoSpaceDN, 51);
	assert.equal(ser.pPrType.Spacing_BeforeLines, 49);
});

test('defaults are true, unset fields stay undefined', () => {
	const p = new CParaPr();
	assert.equal(p.AutoSpaceDE, undefined);
	assert.equal(p.AutoSpaceDN, undefined);
	p.InitDefault(0);
	assert.equal(p.AutoSpaceDE, true);
	assert.equal(p.AutoSpaceDN, true);
});

test('Copy, Merge, Set_FromObject, Is_Equal keep the fields', () => {
	const p = new CParaPr();
	p.AutoSpaceDE = false;
	p.AutoSpaceDN = false;
	const c = p.Copy();
	assert.equal(c.AutoSpaceDE, false);
	assert.equal(c.AutoSpaceDN, false);
	const m = new CParaPr();
	m.Merge(p);
	assert.equal(m.AutoSpaceDE, false);
	assert.equal(m.AutoSpaceDN, false);
	const o = new CParaPr();
	o.Set_FromObject(p);
	assert.equal(o.AutoSpaceDE, false);
	assert.equal(o.AutoSpaceDN, false);
	assert(p.Is_Equal(c));
	c.AutoSpaceDN = true;
	assert(!p.Is_Equal(c));
});

test('history binary round trip keeps the fields', () => {
	const p = new CParaPr();
	p.AutoSpaceDE = false;
	p.AutoSpaceDN = true;
	const memory = new Memory();
	p.Write_ToBinary(memory);
	const q = new CParaPr();
	q.Read_FromBinary(new AC.FT_Stream2(memory.bytes(), memory.size));
	assert.equal(q.AutoSpaceDE, false);
	assert.equal(q.AutoSpaceDN, true);
	assert(p.Is_Equal(q));
});

test('Serialize2 writer emits items 50/51 only when set', () => {
	const plain = new CParaPr();
	const bytesPlain = writePPr(plain);
	assert(!bytesPlain.includes(50) && !bytesPlain.includes(51));
	const p = new CParaPr();
	p.AutoSpaceDE = false;
	p.AutoSpaceDN = false;
	const bytes = writePPr(p);
	const iDE = bytes.indexOf(50), iDN = bytes.indexOf(51);
	assert(iDE > 0 && iDN > iDE);
	assert.equal(bytes[iDE + 1], AC.c_oSerPropLenType.Byte);
	assert.equal(bytes[iDE + 2], 0);
	assert.equal(bytes[iDN + 2], 0);
});

test('Serialize2 reader restores the fields', () => {
	const p = new CParaPr();
	p.AutoSpaceDE = false;
	p.AutoSpaceDN = true;
	const q = readPPr(writePPr(p));
	assert.equal(q.AutoSpaceDE, false);
	assert.equal(q.AutoSpaceDN, true);
	const r = readPPr(writePPr(new CParaPr()));
	assert.equal(r.AutoSpaceDE, undefined);
	assert.equal(r.AutoSpaceDN, undefined);
});
