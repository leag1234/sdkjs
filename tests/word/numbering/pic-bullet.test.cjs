// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Picture bullets in DOCX (w:numPicBullet / w:lvlPicBulletId): binary round trip of the numbering (the new items of
// c_oSerNumTypes), the level model (Copy, IsEqual, history binary), the table of pictures of the numbering and the
// drawing of the bullet in CNum.Measure / CNum.Draw. Run: node --test tests/word/numbering/pic-bullet.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm'), cp = require('node:child_process');
const root = path.resolve(__dirname, '../../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const serial = read('word/Editor/Serialize2.js');
const lvlSrc = read('word/Editor/Numbering/NumberingLvl.js');
const numberingSrc = read('word/Editor/Numbering/Numbering.js');
const numSrc = read('word/Editor/Numbering/Num.js');

function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = to === undefined ? source.length : source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}
// text of a source file as it was in the last commit (to prove that documents without picture bullets are unchanged)
function committed(file) {
	return cp.execFileSync('git', ['show', 'HEAD:' + file], {cwd: root, maxBuffer: 1 << 28}).toString('utf8');
}

// ---- the common serialization code of sdkjs (BinaryCommonWriter, Binary_CommonReader, FT_Stream2) ----
const win = {AscCommon: {}, Asc: {}};
const commonCtx = vm.createContext({
	window: win, AscCommon: win.AscCommon, Asc: win.Asc, console, Uint8Array, TextDecoder,
	AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}}
});
vm.runInContext(read('common/SerializeCommonWordExcel.js'), commonCtx, {filename: 'SerializeCommonWordExcel.js'});
const AC = win.AscCommon;

// simple memory with the interface of the binary writer of sdkjs (little endian)
function Memory() {
	this.data = [];
	this.pos = 0;
	this.size = 0;
	this.GetCurPosition = () => this.pos;
	this.Skip = n => {
		for (let i = 0; i < n; ++i)
			this.WriteByte(0);
	};
	this.Seek = n => {
		this.pos = n;
	};
	this.WriteByte = v => {
		this.data[this.pos++] = v & 0xFF;
		this.size = Math.max(this.size, this.pos);
	};
	this.WriteBool = v => this.WriteByte(v ? 1 : 0);
	this.WriteLong = v => {
		for (let i = 0; i < 4; ++i)
			this.WriteByte((v >> (8 * i)) & 0xFF);
	};
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
function Reader(bytes) {
	return new AC.FT_Stream2(bytes, bytes.length);
}

// ---- c_oSerNumTypes, the writer and the reader of numbering from Serialize2.js ----
const sandbox = {
	CNumberingLvl: null, CAbstractNum: function () {
		this.Lvl = new Array(9);
		this.SetLvl = (i, l) => {
			this.Lvl[i] = l;
		};
		this.SetNsid = () => {};
		this.GetId = () => 'a';
	},
	AscCommon: AC, Asc: win.Asc, AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}},
	c_oSerPropLenType: AC.c_oSerPropLenType, c_oSerConstants: AC.c_oSerConstants,
	BinaryCommonWriter: AC.BinaryCommonWriter, Binary_CommonReader: AC.Binary_CommonReader,
	align_Left: 0, align_Center: 1, align_Right: 2, align_Justify: 3,
	Binary_pPrWriter: function () {
		this.WriteNumFmt = () => {};
		this.Write_pPr = () => {};
	},
	Binary_rPrWriter: function () {
		this.Write_rPr = () => {};
	},
	Binary_pPrReader: function () {
	},
	Binary_rPrReader: function () {
	},
	console
};
function loadSerialize(source) {
	const ctx = vm.createContext(Object.assign({}, sandbox));
	const code = slice(source, 'var c_oSerNumTypes = {', 'var c_oSerOtherTableTypes')
		+ slice(source, 'function BinaryNumberingTableWriter(', 'function BinaryDocumentTableWriter(')
		+ slice(source, 'function Binary_NumberingTableReader(', 'function Binary_HdrFtrTableReader(')
		+ ';this.c_oSerNumTypes = c_oSerNumTypes; this.Writer = BinaryNumberingTableWriter; this.NumReader = Binary_NumberingTableReader;';
	vm.runInContext(code, ctx, {filename: 'Serialize2.js'});
	return ctx;
}
const cur = loadSerialize(serial);


// ---- CNumbering (table of pictures) and CNumberingLvl ----
function loadNumbering() {
	const ctx = vm.createContext({AscWord: {}, AscCommon: AC, Asc: win.Asc});
	const code = slice(numberingSrc, 'function CNumbering()', 'CNumbering.prototype.CopyAllNums');
	vm.runInContext(code.slice(0, code.lastIndexOf('/**')) + ';this.CNumbering = CNumbering;', ctx);
	return ctx.CNumbering;
}
const CNumbering = loadNumbering();

function loadLvl() {
	const ctx = vm.createContext({
		CNumberingLvlLegacy: function () {},
		AscCommon: {align_Left: 0}, Asc: {c_oAscNumberingFormat: {Bullet: 23, Decimal: 1}, c_oAscNumberingSuff: {Tab: 0}},
		CTextPr: function () {
			this.IsEqual = () => true;
			this.Copy = () => new ctx.CTextPr();
			this.WriteToBinary = () => {};
			this.ReadFromBinary = () => {};
		},
		CParaPr: function () {
			this.IsEqual = () => true;
			this.Copy = () => new ctx.CParaPr();
			this.WriteToBinary = () => {};
			this.ReadFromBinary = () => {};
		}
	});
	let code = slice(lvlSrc, 'function CNumberingLvl()', 'CNumberingLvl.prototype.GetJc');
	code = code.slice(0, code.lastIndexOf('/**'));
	code += 'CNumberingLvl.prototype.private_CheckSymbols = function(){};';
	code += 'CNumberingLvl.prototype.SetFormat = function(f){this.Format = f;};';
	code += 'CNumberingLvl.prototype.IsSimilar = function(){return true;};';
	code += 'CNumberingLvl.prototype.private_ReadLvlTextFromBinary = function(){return null;};';
	const method = (m, n) => slice(lvlSrc, m, n);
	code += method('CNumberingLvl.prototype.Copy = ', '/**\n * Выставляем значения по заданному пресету');
	code += method('CNumberingLvl.prototype.IsEqual = ', '/**\n * Заполняем специальный класс для работы с интерфейсом');
	code += method('CNumberingLvl.prototype.WriteToBinary = ', 'CNumberingLvl.prototype.private_ReadLvlTextFromBinary = ');
	code += ';this.CNumberingLvl = CNumberingLvl;';
	vm.runInContext(code, ctx);
	return ctx.CNumberingLvl;
}
const CNumberingLvl = loadLvl();
sandbox.CNumberingLvl = CNumberingLvl;
const cur2 = loadSerialize(serial);
cur.NumReader = cur2.NumReader;

// ----------------------------------------------------------------------------------------------------------------

test('item numbers of sdkjs equal core BinReaderWriterDefines.h and old numbers are not changed', () => {
	const T = cur.c_oSerNumTypes;
	const expected = {
		Nsid: 38, lvl_PicBulletId: 39, NumPicBullets: 40, NumPicBullet: 41, NumPicBullet_Id: 42, NumPicBullet_Image: 43,
		NumPicBullet_Width: 44, NumPicBullet_Height: 45
	};
	for (const k in expected)
		assert.equal(T[k], expected[k], k);
	// the numbers of the existing items are the ones of the previous version
	const old = loadSerialize(committed('word/Editor/Serialize2.js')).c_oSerNumTypes;
	for (const k in old)
		assert.equal(T[k], old[k], k);
	// unique numbers
	const values = Object.values(T);
	assert.equal(new Set(values).size, values.length);

	const core = path.resolve(root, '../core/OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h');
	if (!fs.existsSync(core))
		return;
	const header = slice(fs.readFileSync(core, 'utf8'), 'namespace c_oSerNumTypes', '};}');
	for (const k in T) {
		const m = new RegExp('\\b' + k + '\\s*=\\s*(\\d+)').exec(header);
		assert(m, 'core has ' + k);
		assert.equal(parseInt(m[1], 10), T[k], k);
	}
});

test('table of picture bullets: ids, remap of busy ids, same picture is reused', () => {
	const n = new CNumbering();
	assert.equal(n.HasPicBullets(), false);
	assert.equal(n.GetPicBullet(0), null);
	assert.equal(n.AddPicBullet(0, 'image1.png', 9, 9), 0);
	assert.equal(n.AddPicBullet(0, 'image1.png', 9, 9), 0, 'same picture');
	assert.equal(n.AddPicBullet(0, 'image2.png', 5, 5), 1, 'id 0 is busy');
	assert.equal(n.AddPicBullet(7, 'image3.png', 0, 0), 7);
	assert.equal(n.HasPicBullets(), true);
	assert.equal(n.GetPicBullet(1).Image, 'image2.png');
	assert.equal(n.GetPicBullet(99), null);
});

function writeNumbering(ctx, numbering, lvl) {
	const mem = new Memory();
	const doc = {Numbering: numbering};
	const w = new ctx.Writer(mem, doc, {}, null, {});
	w.bs.WriteItemWithLength(() => w.WriteNumberingContent());
	return mem.bytes();
}
function makeNumbering(withPicture) {
	const n = new CNumbering();
	const lvl = new CNumberingLvl();
	lvl.Format = null; lvl.Jc = null; lvl.LvlText = null; lvl.Suff = null; lvl.TextPr = null; lvl.ParaPr = null;
	lvl.Legacy = null; lvl.IsLgl = null; lvl.Restart = null; lvl.Start = null; lvl.PStyle = null;
	const lvl2 = new CNumberingLvl();
	Object.assign(lvl2, {Format: null, Jc: null, LvlText: null, Suff: null, TextPr: null, ParaPr: null, Legacy: null, IsLgl: null, Restart: null, Start: null, PStyle: null});
	if (withPicture) {
		n.AddPicBullet(3, 'image1.png', 9, 6.5);
		lvl2.PicBulletId = 3;
	}
	const abstract = {GetId: () => 'a1', GetNsid: () => undefined, Lvl: [lvl, lvl2]};
	n.AbstractNum = {a1: abstract};
	n.Num = {};
	return n;
}

test('a numbering without picture bullets gives byte-identical binary data as before', () => {
	const before = loadSerialize(committed('word/Editor/Serialize2.js'));
	const a = writeNumbering(before, makeNumbering(false));
	const b = writeNumbering(cur, makeNumbering(false));
	assert(a.length > 10);
	assert.deepEqual(Array.from(b), Array.from(a));
});

test('the numbering table of picture bullets and the level ids survive write -> read', () => {
	const bytes = writeNumbering(cur, makeNumbering(true));
	const without = writeNumbering(cur, makeNumbering(false));
	assert(bytes.length > without.length);

	// read: only the part of the table of pictures and of the levels is under test
	const stream = Reader(bytes);
	const target = new CNumbering();
	const lvls = [];
	const readResult = {logicDocument: {GetNumbering: () => target}, styleLinks: []};
	const reader = new cur.NumReader({}, readResult, stream);
	const res = reader.Read();
	assert.equal(res, 0);
	assert.equal(reader.m_oANums[0].Lvl[1].PicBulletId, 3);
	assert.equal(reader.m_oANums[0].Lvl[0].PicBulletId, undefined);
	assert.equal(target.GetPicBullet(3).Image, 'image1.png');
	assert.equal(target.GetPicBullet(3).Width, 9);
	assert.equal(target.GetPicBullet(3).Height, 6.5);
});

test('level item lvl_PicBulletId is read with the id of the table (remapped when the id is busy)', () => {
	const mem = new Memory();
	const T = cur.c_oSerNumTypes;
	mem.WriteByte(T.lvl_PicBulletId);
	mem.WriteByte(AC.c_oSerPropLenType.Long);
	mem.WriteLong(5);
	const bytes = mem.bytes();

	const target = new CNumbering();
	target.AddPicBullet(5, 'other.png', 1, 1);
	const readResult = {logicDocument: {GetNumbering: () => target}, styleLinks: []};
	const stream = Reader(bytes);
	const reader = new cur.NumReader({}, readResult, stream);
	reader.m_oPicBulletIds[5] = target.AddPicBullet(5, 'image1.png', 9, 9); // the picture of the file moved to id 6
	assert.equal(reader.m_oPicBulletIds[5], 6);
	const lvl = {};
	const bcr = new AC.Binary_CommonReader(stream);
	bcr.Read2(bytes.length, (t, l) => reader.ReadLevel(t, l, lvl, {nLevelNum: 0}));
	assert.equal(lvl.PicBulletId, 6);
});

test('a level whose picture is not in the table does not write the id', () => {
	const n = makeNumbering(true);
	delete n.PicBullets[3];
	const bytes = writeNumbering(cur, n);
	assert.deepEqual(Array.from(bytes), Array.from(writeNumbering(cur, makeNumbering(false))));
});

test('CNumberingLvl: Copy, IsEqual and the history binary carry the picture bullet id', () => {
	const lvl = new CNumberingLvl();
	assert.equal(lvl.PicBulletId, undefined);
	lvl.PicBulletId = 4;
	const copy = lvl.Copy();
	assert.equal(copy.PicBulletId, 4);
	assert.equal(copy.IsEqual(lvl), true);
	copy.PicBulletId = 5;
	assert.equal(copy.IsEqual(lvl), false);
	copy.PicBulletId = undefined;
	assert.equal(copy.IsEqual(lvl), false);

	for (const id of [undefined, 0, 4]) {
		const l = new CNumberingLvl();
		l.PicBulletId = id;
		const mem = new Memory();
		const w = {
			WriteLong: v => mem.WriteLong(v), WriteBool: v => mem.WriteBool(v), WriteString2: s => mem.WriteString2(s)
		};
		l.WriteToBinary(w);
		const s = Reader(mem.bytes());
		const r = {
			GetLong: () => s.GetLongLE(), GetBool: () => s.GetBool(), GetString2: () => s.GetString2()
		};
		const out = new CNumberingLvl();
		out.ReadFromBinary(r);
		assert.equal(out.PicBulletId, id);
	}
});

// ---- stage 2: drawing ----
function loadNum(images) {
	const ctx = vm.createContext({
		AscCommon: {vertalign_Baseline: 0, vaKSize: 0.65, getFullImageSrc2: s => 'http://doc/media/' + s},
		AscFonts: {ImageLoadStatus: {Loading: 1, Complete: 2}},
		Asc: {editor: {ImageLoader: {map_image_index: images}}}, g_dKoef_pt_to_mm: 25.4 / 72,
		numbering_lvltext_Text: 1, numbering_lvltext_Num: 2,
		g_oTextMeasurer: {SetTextPr() {}}
	});
	ctx.AscWord = {GetFontSlotByTextPr: () => 1, NumBidiDraw: function () {
		this.calls = [];
		this.begin = () => {};
		this.addTextString = s => ctx.drawnText.push(s);
		this.end = () => {};
	}, fontslot_ASCII: 1};
	ctx.drawnText = [];
	// CNum needs only its prototype methods here
	const code = slice(numSrc, 'CNum.prototype.private_GetPicBullet = ', '/**\n * Составляем список всех используемых символов');
	vm.runInContext('function CNum(n){this.Numbering = n;}'
		+ 'String.prototype.getUnicodeIterator = function(){var s=this;return {value:function(){return s.codePointAt(0);}};};'
		+ 'CNum.prototype.private_GetPicBullet = ' + code.slice('CNum.prototype.private_GetPicBullet = '.length) + ';this.CNum = CNum;', ctx);
	return ctx;
}
function lvlWith(picId) {
	return {
		PicBulletId: picId, GetLvlText: () => [{Type: 1, Value: '·'}]
	};
}
function context() {
	const c = {drawn: [], textSet: 0};
	c.drawImage = (...a) => c.drawn.push(a);
	c.SetTextPr = () => c.textSet++;
	c.SetFontSlot = () => {};
	c.GetAscender = () => 3;
	c.MeasureCode = () => ({Width: 1.5});
	return c;
}

test('CNum.Measure / CNum.Draw: a loaded picture is drawn as high as the font, with its aspect ratio', () => {
	const images = {'http://doc/media/image1.png': {Image: {width: 20, height: 10}, Status: 2}};
	const ctx = loadNum(images);
	const numbering = new CNumbering();
	numbering.AddPicBullet(0, 'image1.png', 9, 9);
	const num = new ctx.CNum(numbering);
	num.GetLvl = () => lvlWith(0);
	const textPr = {FontSize: 12, VertAlign: 0};
	const c = context();
	const mm = 12 * 25.4 / 72;
	const measured = num.Measure(c, 0, [1], textPr, null);
	assert(Math.abs(measured.Width - 2 * mm) < 1e-9, 'width = height * 2');
	num.Draw(10, 20, c, 0, [1], textPr, null, false);
	assert.equal(c.drawn.length, 1);
	assert.equal(c.drawn[0][0], 'http://doc/media/image1.png');
	assert.equal(c.drawn[0][1], 10);
	assert(Math.abs(c.drawn[0][2] - (20 - mm)) < 1e-9, 'bottom of the picture on the baseline');
	assert(Math.abs(c.drawn[0][3] - 2 * mm) < 1e-9);
	assert(Math.abs(c.drawn[0][4] - mm) < 1e-9);
	assert.equal(ctx.drawnText.length, 0, 'no text is drawn');
});

test('CNum.Measure / CNum.Draw: a picture that is missing or not loaded falls back to the text bullet', () => {
	for (const images of [{}, {'http://doc/media/image1.png': {Image: null, Status: 2}}, {'http://doc/media/image1.png': {Image: {width: 1, height: 1}, Status: 1}}]) {
		const ctx = loadNum(images);
		const numbering = new CNumbering();
		numbering.AddPicBullet(0, 'image1.png', 9, 9);
		const num = new ctx.CNum(numbering);
		num.GetLvl = () => lvlWith(0);
		num.ApplyTextPrToCodePoint = () => null;
		const c = context();
		const m = num.Measure(c, 0, [1], {FontSize: 12, VertAlign: 0}, null);
		assert.equal(m.Width, 1.5, 'width of the text bullet');
		num.Draw(10, 20, c, 0, [1], {FontSize: 12, VertAlign: 0}, null, false);
		assert.equal(c.drawn.length, 0);
		assert.deepEqual(ctx.drawnText, ['·']);
	}
	// a level with an id that is not in the table, and a level without picture
	for (const id of [42, undefined]) {
		const ctx = loadNum({'http://doc/media/image1.png': {Image: {width: 1, height: 1}, Status: 2}});
		const numbering = new CNumbering();
		numbering.AddPicBullet(0, 'image1.png', 9, 9);
		const num = new ctx.CNum(numbering);
		num.GetLvl = () => lvlWith(id);
		num.ApplyTextPrToCodePoint = () => null;
		const c = context();
		num.Draw(10, 20, c, 0, [1], {FontSize: 12, VertAlign: 0}, null, false);
		assert.equal(c.drawn.length, 0);
		assert.deepEqual(ctx.drawnText, ['·']);
	}
});

test('CNum.Draw: the size of the numbering definition is used when the image has no size', () => {
	const images = {'http://doc/media/image1.png': {Image: {width: 0, height: 0}, Status: 2}};
	const ctx = loadNum(images);
	const numbering = new CNumbering();
	numbering.AddPicBullet(0, 'image1.png', 12, 6);
	const num = new ctx.CNum(numbering);
	num.GetLvl = () => lvlWith(0);
	// the image has no pixel size -> factor 1 -> height = font size - 2 = 10 pt, width = 10 pt * (12/6)
	const mm = 10 * 25.4 / 72;
	assert(Math.abs(num.Measure(context(), 0, [1], {FontSize: 12, VertAlign: 0}, null).Width - 2 * mm) < 1e-9);
});

// ---- stage 3: the Word size rule of the picture bullet ----
// drawn height (pt) = (font size - 2) * shapeHeightPt / (imagePixelHeight * 0.75), clamped to >= 1 pt
function heightPt(ctx, fontSize, shapeH, imgHPx) {
	const num = new ctx.CNum(null);
	const picBullet = {Entry: {Height: shapeH}, ImageHeightPx: imgHPx};
	return num.private_GetPicBulletHeight({FontSize: fontSize}, 1, picBullet) / (25.4 / 72);
}
function ctxWithImage(w, h) {
	const images = {'http://doc/media/image1.png': {Image: {width: w, height: h}, Status: 2}};
	return loadNum(images);
}

test('private_GetPicBulletHeight: the four examples measured in Word', () => {
	const ctx = ctxWithImage(12, 12);
	assert(Math.abs(heightPt(ctx, 12, 9, 12) - 10) < 1e-9, '12 px image, shape 9 pt, 12 pt font -> 10 pt');
	assert(Math.abs(heightPt(ctx, 24, 9, 12) - 22) < 1e-9, 'same at 24 pt font -> 22 pt');
	assert(Math.abs(heightPt(ctx, 12, 9, 32) - 3.75) < 1e-9, '32 px image, shape 9 pt, 12 pt font -> 3.75 pt');
	assert(Math.abs(heightPt(ctx, 12, 36, 12) - 40) < 1e-9, 'shape 36 pt, 12 px image, 12 pt font -> 40 pt');
});

test('private_GetPicBulletHeight: guards', () => {
	const ctx = ctxWithImage(12, 12);
	assert(Math.abs(heightPt(ctx, 12, 0, 12) - 10) < 1e-9, 'shape height 0 -> factor 1');
	assert(Math.abs(heightPt(ctx, 12, undefined, 12) - 10) < 1e-9, 'shape height missing -> factor 1');
	assert(Math.abs(heightPt(ctx, 12, 9, 0) - 10) < 1e-9, 'image pixel height 0 -> factor 1');
	assert(Math.abs(heightPt(ctx, 12, 0.5, 32) - 1) < 1e-9, 'height below 1 pt is clamped to 1 pt');
	assert(Math.abs(heightPt(ctx, 2, 9, 12) - 1) < 1e-9, 'font size - 2 below 1 -> base 1');
});

test('CNum.Draw: the picture bottom sits on the baseline with the Word height', () => {
	const ctx = ctxWithImage(12, 12);
	const numbering = new CNumbering();
	numbering.AddPicBullet(0, 'image1.png', 9, 9);
	const num = new ctx.CNum(numbering);
	num.GetLvl = () => lvlWith(0);
	const c = context();
	const hPt = 10, hMm = hPt * 25.4 / 72;
	num.Draw(10, 20, c, 0, [1], {FontSize: 12, VertAlign: 0}, null, false);
	assert.equal(c.drawn.length, 1);
	assert(Math.abs(c.drawn[0][2] - (20 - hMm)) < 1e-9, 'bottom on the baseline');
	assert(Math.abs(c.drawn[0][3] - hMm) < 1e-9, 'width = height * aspect ratio (square image)');
	assert(Math.abs(c.drawn[0][4] - hMm) < 1e-9, 'height 10 pt');
});
