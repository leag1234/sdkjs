// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Auto-spacing layout pass: classifier, gap constants, and the shaper post-pass.
// Run: node --test tests/word/autospace-layout.test.cjs
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

// ---- classifier ----
const win = {AscWord: {}};
const ctx = vm.createContext({window: win, console});
vm.runInContext(read('word/Editor/Paragraph/Run/AutoSpaceClassifier.js'), ctx, {filename: 'AutoSpaceClassifier.js'});
const AscWord = win.AscWord;

test('classifier follows the measured Word classes', () => {
	const C = AscWord.AUTOSPACE_CLASS;
	assert.equal(AscWord.getAutoSpaceClass('日'.codePointAt(0)), C.EAST_ASIAN);
	assert.equal(AscWord.getAutoSpaceClass('。'.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('、'.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('％'.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('A'.codePointAt(0)), C.LATIN);
	assert.equal(AscWord.getAutoSpaceClass('é'.codePointAt(0)), C.LATIN);
	assert.equal(AscWord.getAutoSpaceClass('α'.codePointAt(0)), C.LATIN);
	assert.equal(AscWord.getAutoSpaceClass('Я'.codePointAt(0)), C.LATIN);
	assert.equal(AscWord.getAutoSpaceClass('ก'.codePointAt(0)), C.LATIN);
	assert.equal(AscWord.getAutoSpaceClass('1'.codePointAt(0)), C.DIGIT);
	assert.equal(AscWord.getAutoSpaceClass('１'.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('('.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('א'.codePointAt(0)), C.OTHER);
	assert.equal(AscWord.getAutoSpaceClass('가'.codePointAt(0)), C.EAST_ASIAN);
});

test('gap constants match the measured rule', () => {
	const C = AscWord.AUTOSPACE_CLASS;
	assert.equal(AscWord.getAutoSpaceGapEm(C.EAST_ASIAN, C.LATIN, 'Times New Roman'), 0.25);
	assert.equal(AscWord.getAutoSpaceGapEm(C.EAST_ASIAN, C.DIGIT, 'Arial'), 0.25);
	assert.equal(AscWord.getAutoSpaceGapEm(C.LATIN, C.EAST_ASIAN, 'Times New Roman'), 0.20);
	assert.equal(AscWord.getAutoSpaceGapEm(C.DIGIT, C.EAST_ASIAN, 'Arial'), 0.22);
	assert.equal(AscWord.getAutoSpaceGapEm(C.EAST_ASIAN, C.EAST_ASIAN, 'Arial'), 0);
	assert.equal(AscWord.getAutoSpaceGapEm(C.LATIN, C.DIGIT, 'Arial'), 0);
	assert.equal(AscWord.getAutoSpaceGapEm(C.EAST_ASIAN, C.OTHER, 'Arial'), 0);
	assert.equal(AscWord.getAutoSpaceGapEm(C.OTHER, C.EAST_ASIAN, 'Arial'), 0);
});

// ---- shaper post-pass ----
function makeItem(codePoint, fontSize, textPr) {
	const item = {
		Value: codePoint,
		Flags: ((fontSize * 64) & 0xffff) << 16,
		Width: 0,
		TempWidth: 0,
		TempGrapheme: null,
		Grapheme: null,
		IsText: () => true,
		IsNBSP: () => false,
		IsPdfText: () => false,
		IsDigit: () => codePoint >= 48 && codePoint <= 57,
		IsLetter: () => (codePoint >= 65 && codePoint <= 90) || (codePoint >= 97 && codePoint <= 122),
		IsSpaceAfter: () => false,
		GetCodePoint: () => codePoint,
		GetFontSize: () => ((item.Flags >> 16) & 0xffff) / 64,
		SetWidth: () => { item.Width = 0; item.Flags &= ~0x100; item.TempWidth = 0; item.autoSpaced = false; },
		SetTemporaryWidth: () => { item.TempWidth = 0; item.Flags |= 0x100; item.Width = 0; item.autoSpaced = false; },
		GetMeasuredWidth: () => 0,
		ResetTemporaryGrapheme: () => {},
		GetGrapheme: () => null,
		GetCodePointType: () => 0,
		SetTemporaryGrapheme: () => {},
		SetTemporaryCodePointType: () => {},
		SetGrapheme: () => {},
		SetMetrics: () => {},
		SetCodePointType: () => {},
		AddAutoSpaceWidth(width) {
			if (item.autoSpaced)
				return;
			item.autoSpaced = true;
			if (item.Flags & 0x100)
				item.TempWidth += width;
			else
				item.Width += width;
		}
	};
	return item;
}

function makeRun(items, fontName) {
	return {
		Content: items,
		GetElementsCount: () => items.length,
		GetElement: i => items[i],
		GetParagraph: () => paragraph,
		Get_CompiledPr: () => ({RFonts: {Ascii: {Name: fontName}, IsEqualSlot: () => true}, IsEqual: () => true, GetFontSize: () => 12, GetSpacing: () => 0, GetLigatures: () => 0}),
		GetParagraph: () => null,
		GetParent: () => null,
		GetParentForm: () => null,
		IsUseAscFont: () => false
	};
}

function makeParagraph(items, paraPr) {
	const run = makeRun(items, 'Times New Roman');
	const paragraph = {
		Content: [run],
		Get_CompiledPr2: () => ({ParaPr: paraPr}),
		CheckRunContent(f, start, end) {
			const s = start ? start.pos : 0;
			const e = end ? end.pos : this.Content.length - 1;
			for (let i = s; i <= e; ++i)
				this.Content[i].CheckRunContent(f, i === s ? start : null, i === e ? end : null, 1, null, true);
		}
	};
	run.CheckRunContent = function(f, s, e) {
		const from = s ? s.inRunPos : 0;
		const to = e ? e.inRunPos : this.Content.length;
		for (let i = from; i < to; ++i)
			f(run, i, i + 1);
	};
	return paragraph;
}

function loadShaper() {
	const win = {AscWord, Asc: {LigaturesType: {None: 0}}, AscFonts: {
		CTextShaper: function () {},
		NO_GRAPHEME: null,
		DEFAULT_TEXTFONTINFO: {Size: 12, Name: 'Times New Roman', Style: 0}
	}, AscCommon: {}, editor: null, console};
	win.AscFonts.CTextShaper.prototype = {};
	win.AscFonts.CTextShaper.prototype.ClearBuffer = function () {};
	win.AscFonts.CTextShaper.prototype.AppendToString = function () {};
	win.AscFonts.CTextShaper.prototype.FlushWord = function () {};
	win.AscFonts.CTextShaper.prototype.IsRtlDirection = function () { return false; };
	win.AscFonts.StringShaper = function () {};
	win.AscFonts.CTextShaper.prototype.HandleRun = function () {};
	const ctx = vm.createContext({window: win, AscWord: Object.assign({}, AscWord, {TEXTWIDTH_DIVIDER: 1}), Asc: win.Asc, AscFonts: win.AscFonts,
		AscCommon: win.AscCommon, editor: {isPdfEditor: () => false}, console,
		IsEqualNullableFloatNumbers: (a, b) => a === b});
	vm.runInContext(read('word/Editor/Paragraph/TextShaper.js'), ctx, {filename: 'TextShaper.js'});
	return win.AscWord.ParagraphTextShaper.constructor;
}

test('shaper post-pass adds the measured gaps to the left item', () => {
	const Shaper = loadShaper();
	const items = [
		makeItem('語'.codePointAt(0), 12),
		makeItem('A'.codePointAt(0), 12),
		makeItem('3'.codePointAt(0), 12),
		makeItem('テ'.codePointAt(0), 12)
	];
	const paragraph = makeParagraph(items, {});
	new Shaper().Shape(paragraph);
	assert.ok(Math.abs(items[0].Width - 12 * 0.25 * 25.4 / 72) < 1e-9); // pt -> mm
	assert.equal(items[1].Width, 0);
	assert.ok(Math.abs(items[2].Width - 12 * 0.20 * 25.4 / 72) < 1e-9); // pt -> mm
	assert.equal(items[3].Width, 0);
});

test('DE and DN are independent', () => {
	const Shaper = loadShaper();
	const make = de => {
		const items = [
			makeItem('語'.codePointAt(0), 12),
			makeItem('A'.codePointAt(0), 12),
			makeItem('3'.codePointAt(0), 12),
			makeItem('テ'.codePointAt(0), 12)
		];
		new Shaper().Shape(makeParagraph(items, {AutoSpaceDE: de, AutoSpaceDN: true}));
		return items;
	};
	assert.equal(make(false)[0].Width, 0);
	assert.ok(Math.abs(make(false)[2].Width - 12 * 0.20 * 25.4 / 72) < 1e-9); // pt -> mm
	assert.ok(Math.abs(make(true)[0].Width - 12 * 0.25 * 25.4 / 72) < 1e-9); // pt -> mm
	assert.ok(Math.abs(make(true)[2].Width - 12 * 0.20 * 25.4 / 72) < 1e-9); // pt -> mm
});

test('spaces, punctuation and full-width forms reset the pair', () => {
	const Shaper = loadShaper();
	const cases = [
		['日', ' ', 'A'],
		['。', 'A'],
		['A', '、'],
		['日', '％']
	];
	for (const chars of cases) {
		const items = chars.map(ch => makeItem(ch.codePointAt(0), 12));
		new Shaper().Shape(makeParagraph(items, {}));
		assert(items.every(item => item.Width === 0), chars.join(''));
	}
});

test('Hangul and Latin use the same rule', () => {
	const Shaper = loadShaper();
	const items = [makeItem('가'.codePointAt(0), 12), makeItem('A'.codePointAt(0), 12)];
	new Shaper().Shape(makeParagraph(items, {}));
	assert.ok(Math.abs(items[0].Width - 12 * 0.25 * 25.4 / 72) < 1e-9); // pt -> mm
});

test('temporary shaping is consistent', () => {
	const Shaper = loadShaper();
	const items = [makeItem('語'.codePointAt(0), 12), makeItem('A'.codePointAt(0), 12)];
	items.forEach(item => { item.Flags |= 0x100; });
	new Shaper().ShapeRange(makeParagraph(items, {}), {pos: 0, inRunPos: 0}, {pos: 0, inRunPos: 2}, true);
	assert.ok(Math.abs(items[0].TempWidth - 12 * 0.25 * 25.4 / 72) < 1e-9); // pt -> mm
});

test('idempotence: reshaping the same text twice does not double the gap', () => {
	const Shaper = loadShaper();
	const items = [makeItem('語'.codePointAt(0), 12), makeItem('A'.codePointAt(0), 12)];
	const paragraph = makeParagraph(items, {});
	const shaper = new Shaper();
	shaper.Shape(paragraph);
	const first = items[0].Width;
	shaper.Shape(paragraph);
	assert.equal(items[0].Width, first);
});
