// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// #39: East Asian document grid (w:docGrid) - model, binary round trip, line snapping and character pitch.
// Run: node --test tests/word/doc-grid.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const recalc = read('word/Editor/Paragraph_Recalculate.js');
const shaper = read('word/Editor/Paragraph/TextShaper.js');
const serialize = read('word/Editor/Serialize2.js');
const styles = read('word/Editor/Styles.js');

function slice(text, from, to) {
	const a = text.indexOf(from);
	assert(a >= 0, from);
	const b = text.indexOf(to, a);
	assert(b > a, to);
	return text.slice(a, b + to.length);
}

const PT = 25.4 / 72;
function loadGrid() {
	const ctx = vm.createContext({AscWord: {}});
	vm.runInContext(read('word/Editor/sections/doc-grid.js'), ctx);
	return ctx.AscWord.SectionDocGrid;
}

class Writer {
	constructor() { this.data = []; }
	GetCurPosition() { return this.data.length; }
	Skip(n) { for (let i = 0; i < n / 4; ++i) this.data.push(0); }
	Seek(pos) { this.pos = pos; }
	WriteLong(v) {
		if (undefined !== this.pos) { this.data[this.pos] = v; this.pos = undefined; return; }
		this.data.push(v);
	}
}
class Reader {
	constructor(data) { this.data = data; this.i = 0; }
	GetLong() { return this.data[this.i++]; }
}

test('model: copy, equality, line and char modes', () => {
	const Grid = loadGrid();
	const g = new Grid(2, 360, 4096);
	assert.ok(g.IsLines() && g.IsChars());
	assert.equal(g.GetLinePitchMM(), 360 * 25.4 / 1440);
	assert.ok(g.Copy().IsEqual(g));
	assert.ok(!g.IsEqual(new Grid(1, 360, 4096)));
	assert.ok(!new Grid(0, 360).IsLines());
	assert.ok(!new Grid(1, 360).IsChars());
	assert.ok(new Grid(3, 360, 0).IsChars());
	assert.equal(new Grid(1, 0).GetLinePitchMM(), 0);
});

test('model: history binary round trip keeps type, pitch and char space', () => {
	const Grid = loadGrid();
	for (const g of [new Grid(1, 360), new Grid(2, 286, 5497), new Grid(2, 457, -2048), new Grid(0)]) {
		const w = new Writer();
		g.Write_ToBinary(w);
		const r = new Grid();
		r.Read_FromBinary(new Reader(w.data));
		assert.ok(r.IsEqual(g), JSON.stringify(g));
	}
});

test('docx binary: sectPr docGrid is written and read with the converter field ids', () => {
	// c_oSerProp_secPrType.docGrid = 14 and c_oSerProp_DocGrid Type/CharSpace/LinePitch = 0/1/2 as in the core
	assert.match(serialize, /lnNumType: 13,\s*docGrid: 14/);
	assert.match(serialize, /var c_oSerProp_DocGrid = \{\s*Type: 0,\s*CharSpace: 1,\s*LinePitch: 2\s*\}/);
	const write = slice(serialize, 'this.WriteDocGrid = function(docGrid)', '\n\t};');
	assert.match(write, /c_oSerProp_DocGrid\.Type[\s\S]*WriteByte\(docGrid\.Type\)/);
	assert.match(write, /c_oSerProp_DocGrid\.CharSpace[\s\S]*WriteLong\(docGrid\.CharSpace\)/);
	assert.match(write, /c_oSerProp_DocGrid\.LinePitch[\s\S]*WriteLong\(docGrid\.LinePitch\)/);
	const readFn = slice(serialize, 'this.Read_docGrid = function(type, length, docGrid)', '\n\t}');
	assert.match(readFn, /GetUChar/);
	assert.match(readFn, /CharSpace = this\.stream\.GetLongLE/);
	assert.match(readFn, /LinePitch = this\.stream\.GetLongLE/);
	assert.match(serialize, /sectPr\.DocGrid\)\s*this\.bs\.WriteItem\(c_oSerProp_secPrType\.docGrid/);
	assert.match(serialize, /oSectPr\.SetDocGrid\(docGrid\)/);
});

test('docx binary: snapToGrid of paragraphs and runs is read and written', () => {
	assert.match(serialize, /case c_oSerProp_pPrType\.SnapToGrid:\s*pPr\.SnapToGrid = this\.stream\.GetBool\(\)/);
	assert.match(serialize, /case c_oSerProp_rPrType\.SnapToGrid:\s*rPr\.SnapToGrid = this\.stream\.GetBool\(\)/);
	assert.match(serialize, /WriteByte\(c_oSerProp_pPrType\.SnapToGrid\)/);
	assert.match(serialize, /WriteByte\(c_oSerProp_rPrType\.SnapToGrid\)/);
	// the default of a compiled paragraph and run is "snap"
	assert.match(styles, /this\.SnapToGrid\s*=\s*true;/);
});

function loadSnap(grid, flags) {
	const ctx = vm.createContext({Asc: {linerule_Auto: 1, linerule_Exact: 2, linerule_AtLeast: 0}, Paragraph: function() {}});
	const src = slice(recalc, 'Paragraph.prototype.private_SnapLineToDocGrid = function', '\n};\n') + '\n' +
		slice(recalc, 'Paragraph.prototype.private_IsDocGridSnapInTable = function', '\n};\n');
	vm.runInContext(src, ctx);
	const Grid = loadGrid();
	const para = Object.create(ctx.Paragraph.prototype);
	para.GetLogicDocument = () => ({GetDocumentSettings: () => ({CompatFlags: flags || null})});
	const make = (asc, desc, gap) => ({Ascent: asc, Descent: desc, LineGap: gap || 0});
	return {
		Grid,
		run(metrics, paraPr, inTable) {
			para.Lines = [{Metrics: metrics}];
			const prs = {GetSectPr: () => (grid ? {GetDocGrid: () => grid} : null), IsInTable: () => !!inTable};
			para.private_SnapLineToDocGrid(0, prs, Object.assign({SnapToGrid: true, Spacing: {LineRule: 1, Line: 1}}, paraPr));
			return metrics;
		},
		make
	};
}
const height = m => m.Ascent + m.Descent + m.LineGap;

test('lines: a single line becomes a whole number of pitches, the text stays centered', () => {
	const t = loadSnap(null);
	const Grid = t.Grid;
	const s = loadSnap(new Grid(1, 360));
	const pitch = 18 * PT;
	// 12 pt MS Mincho: 1.30 em = 15.6 pt -> 1 pitch (Word: 18 pt)
	let m = s.run(s.make(12 * PT, 3.6 * PT), {});
	assert.ok(Math.abs(height(m) - pitch) < 1e-9);
	assert.ok(Math.abs(m.Ascent - (12 + 1.2) * PT) < 1e-9);
	assert.ok(Math.abs(m.LineGap - 1.2 * PT) < 1e-9);
	// 16 pt: 20.8 pt -> 2 pitches; YaHei 10.5 pt: 18.14 pt -> 2 pitches
	assert.ok(Math.abs(height(s.run(s.make(16 * PT, 4.8 * PT), {})) - 2 * pitch) < 1e-9);
	assert.ok(Math.abs(height(s.run(s.make(14 * PT, 4.14 * PT), {})) - 2 * pitch) < 1e-9);
	// exactly one pitch stays one pitch
	assert.ok(Math.abs(height(s.run(s.make(14 * PT, 4 * PT), {})) - pitch) < 1e-9);
});

test('lines: multiple spacing m gives at least m pitches and at least the single height, m below 1 does not shrink', () => {
	const s = loadSnap(new (loadSnap(null).Grid)(1, 360));
	const pitch = 18 * PT;
	const auto = line => ({Spacing: {LineRule: 1, Line: line}});
	// measured with Word 16: 10.5 / 12 pt -> 18, 27, 36, 54 pt for m = 1, 1.5, 2, 3
	for (const [m, expected] of [[1.5, 27], [2, 36], [3, 54], [1.2, 21.6], [0.8, 18]])
		assert.ok(Math.abs(height(s.run(s.make(12 * PT, 3.6 * PT), auto(m))) - expected * PT) < 1e-9, String(m));
	// 24 pt text (single height 36 pt): m = 1.5 and 2 stay at 36 pt, m = 3 gives 54 pt
	for (const [m, expected] of [[1, 36], [1.5, 36], [2, 36], [3, 54]])
		assert.ok(Math.abs(height(s.run(s.make(24 * PT, 7.2 * PT), auto(m))) - expected * PT) < 1e-9, String(m));
	assert.ok(pitch > 0);
});

test('lines: exact and at-least spacing, snapToGrid=false, no grid and type default are not snapped', () => {
	const Grid = loadSnap(null).Grid;
	const s = loadSnap(new Grid(1, 360));
	for (const rule of [2, 0]) {
		const m = s.run(s.make(12 * PT, 3.6 * PT, 1), {Spacing: {LineRule: rule, Line: 20 * PT}});
		assert.ok(Math.abs(height(m) - (15.6 * PT + 1)) < 1e-9);
	}
	let m = s.run(s.make(12 * PT, 3.6 * PT), {SnapToGrid: false});
	assert.ok(Math.abs(height(m) - 15.6 * PT) < 1e-9);
	m = loadSnap(null).run(s.make(12 * PT, 3.6 * PT), {});
	assert.ok(Math.abs(height(m) - 15.6 * PT) < 1e-9);
	m = loadSnap(new Grid(0, 360)).run(s.make(12 * PT, 3.6 * PT), {});
	assert.ok(Math.abs(height(m) - 15.6 * PT) < 1e-9);
});

test('lines: table cells snap only with adjustLineHeightInTable, doNotSnapToGridInCell changes nothing', () => {
	const Grid = loadSnap(null).Grid;
	const grid = new Grid(1, 360);
	const natural = 15.6 * PT;
	for (const [flags, snapped] of [[null, false], [[0, 0, 0], false], [[0, 1 << 2, 0], true], [[0, 1 << 10, 0], false], [[0, (1 << 2) | (1 << 10), 0], true]]) {
		const s = loadSnap(grid, flags);
		const m = s.run(s.make(12 * PT, 3.6 * PT), {}, true);
		assert.ok(Math.abs(height(m) - (snapped ? 18 * PT : natural)) < 1e-9, JSON.stringify(flags));
	}
});

function loadSpacing(grid) {
	const ctx = vm.createContext({});
	const fn = slice(shaper, 'CParagraphTextShaper.prototype.private_GetDocGridSpacing = function', '\n\t};');
	vm.runInContext('function CParagraphTextShaper() {}\n' + fn, ctx);
	const shaperObj = new ctx.CParagraphTextShaper();
	const paragraph = {Get_SectPr: () => (grid ? {GetDocGrid: () => grid} : null)};
	shaperObj.Paragraph = paragraph;
	shaperObj.GridParagraph = null;
	return shaperObj;
}

test('chars: linesAndChars adds charSpace/4096 pt to every character, except in runs with snapToGrid=false', () => {
	const Grid = loadGrid();
	const s = loadSpacing(new Grid(2, 457, 4096));
	assert.ok(Math.abs(s.private_GetDocGridSpacing({}) - PT) < 1e-12);
	assert.equal(s.private_GetDocGridSpacing({SnapToGrid: false}), 0);
	assert.ok(Math.abs(s.private_GetDocGridSpacing({SnapToGrid: true}) - PT) < 1e-12);
	assert.ok(Math.abs(loadSpacing(new Grid(2, 360, -2048)).private_GetDocGridSpacing({}) + 0.5 * PT) < 1e-12);
	assert.ok(Math.abs(loadSpacing(new Grid(2, 360, 5734)).private_GetDocGridSpacing({}) - 5734 / 4096 * PT) < 1e-12);
});

test('chars: type lines, no grid and no charSpace add no spacing', () => {
	const Grid = loadGrid();
	assert.equal(loadSpacing(new Grid(1, 360, 4096)).private_GetDocGridSpacing({}), 0);
	assert.equal(loadSpacing(new Grid(2, 360)).private_GetDocGridSpacing({}), 0);
	assert.equal(loadSpacing(null).private_GetDocGridSpacing({}), 0);
});
