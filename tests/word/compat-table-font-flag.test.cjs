// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// #66: the compatibility setting overrideTableStyleFontSizeAndJustification is kept when a document is saved.
// Word reads a missing setting as false before Word 2013 (compatibility mode 15) and then draws the table text in the
// docDefaults size instead of the Normal style size: writing a fixed "1" changed how Word shows a saved legacy file.
// Run: node --test tests/word/compat-table-font-flag.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const root = path.resolve(__dirname, '../..');
const ser = fs.readFileSync(path.join(root, 'word/Editor/Serialize2.js'), 'utf8');
const settings = fs.readFileSync(path.join(root, 'word/Editor/DocumentSettings.js'), 'utf8');

test('the document settings hold the flag, undefined for a new document', () => {
	assert.match(settings, /this\.OverrideTableStyleFontSizeAndJustification\s*=\s*undefined;/);
});

test('the reader parses the setting value', () => {
	assert.ok(ser.includes('"overrideTableStyleFontSizeAndJustification" === compat.name'));
	const m = /Settings\.OverrideTableStyleFontSizeAndJustification = (\/[^;]+);/.exec(ser);
	assert(m, 'assignment found');
	const re = eval(m[1].replace(/\.test\(.*$/, ''));
	for (const v of ['1', 'true', 'on', 'True', ' 1 '])
		assert.ok(re.test(String(v).trim()), v);
	for (const v of ['0', 'false', 'off', ''])
		assert.ok(!re.test(String(v).trim()), v);
});

test('a file without the setting gets the default of its compatibility mode', () => {
	const i = ser.indexOf('if (undefined === Settings.OverrideTableStyleFontSizeAndJustification)');
	assert(i > 0);
	assert.ok(ser.slice(i, i + 250).includes('Settings.CompatibilityMode >= AscCommon.document_compatibility_mode_Word15'));
});

test('the writer writes the stored value and 1 for a new document', () => {
	const m = /WriteCompatSetting\("overrideTableStyleFontSizeAndJustification", "http:\/\/schemas.microsoft.com\/office\/word", (.*?)\);\}\);/.exec(ser);
	assert(m, 'writer call found');
	const f = new Function('oThis', 'return ' + m[1] + ';');
	const make = value => ({Document: {GetDocumentSettings: () => ({OverrideTableStyleFontSizeAndJustification: value})}});
	assert.equal(f(make(false)), '0');
	assert.equal(f(make(true)), '1');
	assert.equal(f(make(undefined)), '1');
});
