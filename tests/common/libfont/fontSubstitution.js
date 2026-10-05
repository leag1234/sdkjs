/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

// Font substitution (map.js, CApplicationFonts.GetSubstituteName) and symbol-font
// code-to-Unicode replacement (CheckReplaceGlyphsMap / GetReplaceGlyph).
$(function ()
{
	const oApp = AscFonts.g_fontApplication;
	const oDict = oApp.g_fontDictionary;
	if (!oDict.FONTS_DICT_ASCII_NAMES_COUNT)
		oDict.Init();

	// Runs fn with exactly the given installed font families.
	function withInstalled(aNames, fn)
	{
		const oSaved = AscFonts.g_map_font_index;
		const oMap = {};
		aNames.forEach(function (sName, i) { oMap[sName] = i; });
		AscFonts.g_map_font_index = oMap;
		try { fn(); }
		finally { AscFonts.g_map_font_index = oSaved; }
	}

	QUnit.module("Font substitution");

	QUnit.test("installed font always wins over its clone", function (assert)
	{
		withInstalled(["Calibri", "Carlito", "Andika"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Calibri"), "Calibri", "real Calibri");
			assert.strictEqual(oApp.GetSubstituteName("Andika"), "Andika", "real school font");
		});
	});

	QUnit.test("metric-compatible clone when the font is missing", function (assert)
	{
		withInstalled(["Carlito", "Caladea", "Liberation Sans", "Gelasio Metric", "Gelasio", "Symbola"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Calibri"), "Carlito");
			assert.strictEqual(oApp.GetSubstituteName("Cambria"), "Caladea");
			assert.strictEqual(oApp.GetSubstituteName("Arial"), "Liberation Sans");
			assert.strictEqual(oApp.GetSubstituteName("Georgia"), "Gelasio Metric", "first installed substitute in order");
			assert.strictEqual(oApp.GetSubstituteName("Wingdings"), "Symbola");
		});
		withInstalled(["Elixir Symbols", "Symbola"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Wingdings"), "Elixir Symbols", "symbol font with Wingdings metrics first");
			assert.strictEqual(oApp.GetSubstituteName("Webdings"), "Elixir Symbols");
		});
		withInstalled(["Gelasio"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Georgia"), "Gelasio", "next substitute in order");
		});
	});

	QUnit.test("school fonts that cannot be shipped use their measured substitute", function (assert)
	{
		withInstalled(["Carlito", "URW Gothic Ecole Metric"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Script cole"), "URW Gothic Ecole Metric", "Script École");
			assert.strictEqual(oApp.GetSubstituteName("KG Primary Penmanship"), "KG Primary Penmanship", "substitute not installed: name kept");
		});
		withInstalled(["Script cole", "URW Gothic Ecole Metric"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Script cole"), "Script cole", "installed school font wins");
		});
	});

	QUnit.test("unknown font falls back to Word's default font (Calibri, then its clone)", function (assert)
	{
		withInstalled(["Carlito", "Liberation Sans"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Andika"), "Carlito");
			assert.strictEqual(oApp.GetSubstituteName("Cursive standard"), "Carlito");
			assert.strictEqual(oApp.GetSubstituteName("Marianne"), "Carlito");
		});
		withInstalled(["Calibri"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Belle Allure CM"), "Calibri");
		});
	});

	QUnit.test("alternate name of the document font (w:altName), as Word", function (assert)
	{
		oApp.SetDocumentAltNames({"Marianne" : "Times New Roman", "Script cole" : "Courier New", "Loop A" : "Loop B", "Loop B" : "Loop A", "Bradley Hand ITC" : "Viner Hand ITC"});
		try
		{
			withInstalled(["Carlito", "Liberation Serif", "Liberation Mono", "URW Gothic Ecole Metric"], function ()
			{
				assert.strictEqual(oApp.GetSubstituteName("Marianne"), "Liberation Serif", "alternate name, then its clone");
				assert.strictEqual(oApp.GetSubstituteName("Script cole"), "URW Gothic Ecole Metric", "measured substitute before the alternate name");
				assert.strictEqual(oApp.GetSubstituteName("Bradley Hand ITC"), "Bradley Hand ITC", "alternate name not installed: the font keeps its own dictionary data");
				assert.strictEqual(oApp.GetSubstituteName("Loop A"), "Carlito", "alternate names not installed: unknown font rule");
			});
			withInstalled(["Times New Roman", "Marianne"], function ()
			{
				assert.strictEqual(oApp.GetSubstituteName("Marianne"), "Marianne", "installed font wins");
			});
		}
		finally
		{
			oApp.SetDocumentAltNames({});
		}
	});

	QUnit.test("known but missing font without substitute keeps the dictionary matching", function (assert)
	{
		withInstalled(["Carlito", "Liberation Sans"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Tahoma"), "Tahoma", "selection by panose/metrics as before");
			assert.strictEqual(oApp.GetSubstituteName("Book Antiqua"), "Book Antiqua", "listed substitute not installed");
		});
	});

	QUnit.module("Symbol fonts");

	function replace(sFont, sDst, nCode)
	{
		const oDst = {Name : sDst};
		const oReplace = oApp.CheckReplaceGlyphsMap(sFont, oDst);
		return oReplace ? oApp.GetReplaceGlyph(nCode, oReplace) : nCode;
	}

	QUnit.test("Wingdings codes become Unicode symbols on a substitute", function (assert)
	{
		assert.strictEqual(replace("Wingdings", "Symbola", 0xF04A), 0x263A, "w:sym F04A: smiley");
		assert.strictEqual(replace("Wingdings", "Symbola", 0x4B), 0x1F610, "code 4B: neutral face");
		assert.strictEqual(replace("Wingdings", "Symbola", 0x4C), 0x2639, "code 4C: frowning face");
		assert.strictEqual(replace("Wingdings", "Symbola", 0xFC), 0x2713, "check mark");
		assert.strictEqual(replace("Wingdings", "Symbola", 0x20), 0x20, "space unchanged");
		assert.strictEqual(replace("Wingdings", "Symbola", 0x2192), 0x2192, "Unicode text unchanged");
		assert.strictEqual(replace("Wingdings", "OpenSymbol", 0x6C), 0xE12C, "OpenSymbol keeps its own bullet glyph");
		assert.strictEqual(replace("Wingdings", "OpenSymbol", 0x4A), 0x263A, "other codes mapped on OpenSymbol too");
		assert.strictEqual(replace("Wingdings 2", "Symbola", 0x50), 0x1F5F8, "Wingdings 2 check mark");
		assert.strictEqual(replace("Webdings", "Symbola", 0x61), 0x2714, "Webdings heavy check mark");
	});

	QUnit.test("installed symbol font is left alone", function (assert)
	{
		assert.strictEqual(replace("Wingdings", "Wingdings", 0xF04A), 0xF04A);
		assert.strictEqual(replace("Symbol", "Symbol", 0xF061), 0xF061);
	});

	QUnit.test("Symbol codes become Unicode", function (assert)
	{
		assert.strictEqual(replace("Symbol", "Liberation Serif", 0xF061), 0x03B1, "alpha");
		assert.strictEqual(replace("Symbol", "Liberation Serif", 0xB7), 0x2022, "bullet");
		assert.strictEqual(replace("Symbol", "OpenSymbol", 0xB7), 0xE12C, "OpenSymbol bullet glyph");
	});
});
