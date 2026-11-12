// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Tests for EoSelectStylesToWrite: gallery filtering, closure, defaults, fallback.
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

const win = {AscCommon: {}, Asc: {}};
const commonCtx = vm.createContext({
	window: win, AscCommon: win.ASC, Asc: win.Asc, console, Uint8Array, TextDecoder,
	AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}}
});
vm.runInContext(read('common/SerializeCommonWordExcel.js'), commonCtx, {filename: 'SerializeCommonWordExcel.js'});
const AC = win.AscCommon;

const sandbox = {
	AscCommon: AC, Asc: win.Asc, AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}},
	c_oSerPropLenType: AC.cSerPropLenType || {}, c_oSerConstants: AC.c_oSerConstants || {},
	BinaryCommonWriter: AC.BinaryCommonWriter, Binary_CommonReader: AC.Binary_CommonReader,
	align_Left: 0, align_Center: 1, align_Right: 2, align_Justify: 3,
	Binary_rPrWriter: function () { this.Write_rPr = () => {}; },
	Binary_rPrReader: function () {},
	WriteTrackRevision: () => {},
	console
};
function loadSelector() {
	const ctx = vm.createContext(Object.assign({}, sandbox));
	ctx.AscWord = {};
	const src = read('word/Editor/Serialize2.js');
	const code = slice(src, 'AscWord.EoSelectStylesToWrite = function', 'var c_oSerTableTypes = {')
		+ ';this.EoSelectStylesToWrite = AscWord.EoSelectStylesToWrite;';
	vm.runInContext(code, ctx, {filename: 'Serialize2.js'});
	return ctx.EoSelectStylesToWrite;
}
const select = loadSelector();

function makeStyle(id, marked, basedOn, next, link, name) {
	return {
		id, name: name || '', EoBuiltinGallery: !!marked,
		GetBasedOn: () => basedOn || null, GetNext: () => next || null, GetLink: () => link || null
	};
}
function makeStyles(list, defaults) {
	const styles = {Style: {}};
	list.forEach(s => styles.Style[s.id] = makeStyle(s.id, s.marked, s.basedOn, s.next, s.link, s.name));
	styles.Default = defaults || {};
	styles.GetStyleIdByName = name => {
		const found = list.find(s => s.name === name);
		return found ? found.id : null;
	};
	return styles;
}

test('gallery style unused is skipped', () => {
	const styles = makeStyles([
		{id: 'normal', name: 'Normal', marked: true},
		{id: 'gallery', name: 'Gallery', marked: true}
	], {Paragraph: 'normal', Character: 'char', Table: 'table', Numbering: 'num'});
	const result = select(styles, null);
	assert('normal' in result);
	assert(!('gallery' in result));
});

test('gallery style used by table is kept with based-on chain', () => {
	const styles = makeStyles([
		{id: 'normal', name: 'Normal', marked: true},
		{id: 'base', name: 'Base', marked: true},
		{id: 'tablestyle', name: 'TableStyle', marked: true, basedOn: 'base'}
	], {Paragraph: 'normal', Character: 'char', Table: 'table', Numbering: 'num'});
	const doc = {
		GetAllParagraphs: () => [],
		GetAllTables: () => [{TableStyle: 'tablestyle'}],
		Get_Number: () => null
	};
	const result = select(styles, doc);
	assert('tablestyle' in result);
	assert('base' in result);
});

test('unmarked style unused is kept', () => {
	const styles = makeStyles([
		{id: 'normal', name: 'Normal', marked: true},
		{id: 'user', name: 'User', marked: false}
	], {Paragraph: 'normal', Character: 'char', Table: 'table', Numbering: 'num'});
	const result = select(styles, null);
	assert('user' in result);
});

test('defaults are kept', () => {
	const styles = makeStyles([
		{id: 'normal', name: 'Normal', marked: true},
		{id: 'dpf', name: 'Default Paragraph Font', marked: true},
		{id: 'nolist', name: 'No List', marked: true}
	], {Paragraph: 'normal', Character: 'dpf', Table: 'table', Numbering: 'nolist'});
	const result = select(styles, null);
	assert('normal' in result);
	assert('dpf' in result);
	assert('nolist' in result);
});

test('exception falls back to all styles', () => {
	const styles = makeStyles([
		{id: 'normal', name: 'Normal', marked: true},
		{id: 'gallery', name: 'Gallery', marked: true}
	], {Paragraph: 'normal', Character: 'char', Table: 'table', Numbering: 'num'});
	const doc = {
		GetAllParagraphs: () => { throw new Error('test'); },
		GetAllTables: () => [],
		Get_Number: () => null
	};
	const result = select(styles, doc);
	assert('normal' in result);
	assert('gallery' in result);
});
