/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
'use strict';
// GitHub issue #73: an empty paragraph whose paragraph mark has w:sz=0 (FontSize 0)
// must still get a non-zero line height. Word clamps a font size below 0.5 pt to
// 0.5 pt for layout/drawing metrics only; the stored FontSize must stay 0 so the
// document still saves w:sz="0".
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const root = path.join(__dirname, '../..');
const styles = fs.readFileSync(path.join(root, 'word/Editor/Styles.js'), 'utf8');
const fromToJSON = fs.readFileSync(path.join(root, 'word/fromToJSON.js'), 'utf8');

// Extract the real CTextPr.prototype.GetFontInfo from Styles.js and run it with
// minimal stubs (RFonts slots, AscFonts.CTextFontInfo, AscWord font slots).
function loadGetFontInfo() {
	const start = styles.indexOf('CTextPr.prototype.GetFontInfo = function');
	assert(start >= 0, 'GetFontInfo found in Styles.js');
	const end = styles.indexOf('\n};', start);
	assert(end > start);
	const context = vm.createContext({
		CTextPr: function() {},
		AscWord: {fontslot_ASCII: 1, fontslot_CS: 2, fontslot_EastAsia: 4, fontslot_HAnsi: 8},
		AscFonts: {CTextFontInfo: function(name, style, size) { this.Name = name; this.Style = style; this.Size = size; }}
	});
	vm.runInContext(styles.slice(start, end + 3) + '\nthis.GetFontInfo = CTextPr.prototype.GetFontInfo;', context);
	return context.GetFontInfo;
}

function makeTextPr(fontSize, fontSizeCS) {
	return {
		FontSize: fontSize, FontSizeCS: fontSizeCS,
		Bold: false, Italic: false, BoldCS: false, ItalicCS: false,
		RFonts: {
			Ascii: {Name: 'Open Sans'}, HAnsi: {Name: 'Open Sans'},
			CS: {Name: 'Open Sans'}, EastAsia: {Name: 'Open Sans'}
		}
	};
}

test('GetFontInfo clamps a layout font size below 0.5 pt to 0.5 pt, stored FontSize untouched', () => {
	const getFontInfo = loadGetFontInfo();
	const tp = makeTextPr(0, 0);
	const info = getFontInfo.call(tp, 1 /* fontslot_ASCII */);
	assert.equal(info.Size, 0.5, 'size 0 is clamped to 0.5 for metrics');
	assert.equal(tp.FontSize, 0, 'stored FontSize is NOT modified');
	assert.equal(getFontInfo.call(tp, 2 /* fontslot_CS */).Size, 0.5, 'FontSizeCS 0 clamped too');
});

test('GetFontInfo leaves sizes >= 0.5 pt unchanged', () => {
	const getFontInfo = loadGetFontInfo();
	for (const size of [0.5, 1, 11, 22.75]) {
		const tp = makeTextPr(size, size);
		assert.equal(getFontInfo.call(tp, 1).Size, size);
		assert.equal(getFontInfo.call(tp, 2).Size, size);
	}
});

test('empty paragraph with FontSize 0 gets a positive line height (issue #73 scenario)', () => {
	const getFontInfo = loadGetFontInfo();
	// Open Sans ascent+descent factor ~1.362 (per issue report); line rule auto 1.15.
	const factor = 1.362, lineRule = 1.15, spacingAfterPt = 10;
	const tp = makeTextPr(0, 0);
	const linePt = getFontInfo.call(tp, 1).Size * factor * lineRule;
	const paraPt = linePt + spacingAfterPt;
	assert.ok(linePt > 0, 'line height is positive, not 0');
	assert.ok(paraPt > spacingAfterPt, 'paragraph height exceeds the 10 pt spacing');
	assert.ok(Math.abs(paraPt - (10 + 0.5 * factor * lineRule)) < 1e-9, 'matches Word ~10.78 pt');
	// control: FontSize 11 is unaffected by the clamp
	const tp11 = makeTextPr(11, 11);
	assert.equal(getFontInfo.call(tp11, 1).Size, 11);
});

test('w:sz=0 still reads into FontSize 0 (stored value unchanged, saves back as sz=0)', () => {
	assert.ok(fromToJSON.includes('this.FontSize = oParsedJson["sz"] / 2.0;'), 'read path keeps raw value');
	assert.ok(fromToJSON.includes('oResult["sz"] = 2.0 * this.FontSize;'), 'write path keeps raw value');
});
