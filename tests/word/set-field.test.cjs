/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
'use strict';
// SET field (Word bookmark definition): the instruction `SET name "value"` must be
// recognised as fieldtype_SET, and the field's cached result must be hidden from
// layout/drawing (CComplexField.IsHidden / ParaField.IsHiddenResult) while staying
// in the model. Other field types (ASK, REF, PAGE) keep their types and stay visible.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const root = path.join(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');

// Load the real field-type constants, instruction base class and instruction parser
// into one vm context, the same way the concatenated sdk bundle would see them.
function loadParserContext() {
	const context = vm.createContext({
		window: {},
		AscWord: undefined,
		AscCommonWord: {},
		AscCommon: {}
	});
	context.window = context; // the files do (function(window){...})(window)
	vm.runInContext(read('word/Editor/Paragraph/ComplexFields/types.js'), context);
	vm.runInContext(read('word/Editor/Paragraph/ComplexFields/instruction-base.js'), context);
	vm.runInContext(read('word/Editor/Paragraph/ComplexFields/ref.js'), context);
	vm.runInContext(read('word/Editor/Paragraph/ComplexFieldInstruction.js'), context);
	return context;
}

function parse(ctx, instr) {
	return vm.runInContext(
		'new CFieldInstructionParser().GetInstructionClass(' + JSON.stringify(instr) + ')',
		ctx
	);
}

// Extract a single prototype method from a source file and return it as a callable.
function extractMethod(file, marker, nextMarker) {
	const src = read(file);
	const start = src.indexOf(marker);
	assert(start >= 0, marker + ' found in ' + file);
	const end = src.indexOf(nextMarker, start + marker.length);
	assert(end > start, nextMarker + ' found after ' + marker);
	return src.slice(start, end + 3); // include closing "};"
}

test('SET instruction parses as fieldtype_SET with bookmark name and value', () => {
	const ctx = loadParserContext();
	const AscWord = ctx.AscWord;
	assert.equal(AscWord.fieldtype_SET, 0x0018);

	const instr = parse(ctx, ' SET ref "2001" ');
	assert.ok(instr, 'instruction parsed');
	assert.equal(instr.GetType(), AscWord.fieldtype_SET);
	assert.equal(instr.Type, AscWord.fieldtype_SET);
	assert.equal(instr.GetBookmarkName(), 'ref');
	assert.equal(instr.GetValue(), '2001');
});

test('SET parsing is case-insensitive on the first token', () => {
	const ctx = loadParserContext();
	const instr = parse(ctx, 'set ref "2001"');
	assert.equal(instr.GetType(), ctx.AscWord.fieldtype_SET);
	assert.equal(instr.GetBookmarkName(), 'ref');
	assert.equal(instr.GetValue(), '2001');
});

test('ASK, REF and PAGE keep their own types and are not SET', () => {
	const ctx = loadParserContext();
	const AscWord = ctx.AscWord;

	const ask = parse(ctx, ' ASC ref "prompt" ');
	assert.equal(ask.GetType(), AscWord.fieldtype_ASK);
	assert.notEqual(ask.GetType(), AscWord.fieldtype_SET);

	const ref = parse(ctx, ' REF ref \\h ');
	assert.equal(ref.GetType(), AscWord.fieldtype_REF);
	assert.notEqual(ref.GetType(), AscWord.fieldtype_SET);

	const page = parse(ctx, ' PAGE ');
	assert.equal(page.GetType(), AscWord.fieldtype_PAGE);
	assert.notEqual(page.GetType(), AscWord.fieldtype_SET);
});

test('CComplexField.IsHidden (the layout predicate) is true for SET and ASK, false for REF', () => {
	const src = extractMethod(
		'word/Editor/Paragraph/ComplexField.js',
		'CComplexField.prototype.IsHidden = function()',
		'CComplexField.prototype.RemoveFieldWrap = function()'
	);
	const ctx = loadParserContext();
	vm.runInContext(src + '\nthis.IsHidden = CComplexField.prototype.IsHidden;', Object.assign(ctx, {CComplexField: function() {}, CCo: function() {}}));
	const isHidden = ctx.IsHidden;

	const mkField = type => ({
		SeparateChar: {IsHiddenValue: () => false},
		BeginChar: {},
		GetInstruction: () => ({GetType: () => type})
	});

	assert.equal(isHidden.call(mkField(ctx.AscWord.fieldtype_SET)), true, 'SET result hidden');
	assert.equal(isHidden.call(mkField(ctx.AscWord.fieldtype_ASK)), true, 'ASK result hidden');
	assert.equal(isHidden.call(mkField(ctx.AscWord.fieldtype_REF)), false, 'REF result shown');
	assert.equal(isHidden.call(mkField(ctx.AscWord.fieldtype_PAGE)), false, 'PAGE result shown');
});

test('ParaField.IsHiddenResult is true only for fieldtype_SET', () => {
	const src = extractMethod(
		'word/Editor/Field.js',
		'ParaField.prototype.IsHiddenResult = function()',
		'ParaField.prototype.Map_MailMerge = function(_Value)'
	);
	const ctx = loadParserContext();
	vm.runInContext(src + '\nthis.IsHiddenResult = ParaField.prototype.IsHiddenResult;', Object.assign(ctx, {ParaField: function() {}, Par: function() {}}));
	const isHiddenResult = ctx.IsHiddenResult;

	assert.equal(isHiddenResult.call({FieldType: ctx.AscWord.fieldtype_SET}), true);
	assert.equal(isHiddenResult.call({FieldType: ctx.AscWord.fieldtype_REF}), false);
	assert.equal(isHiddenResult.call({FieldType: ctx.AscWord.fieldtype_PAGE}), false);
});
