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

// Font substitution (map.js, CApplicationFonts.GetSubstituteName).
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
		withInstalled(["Carlito", "Caladea", "Liberation Sans", "Gelasio", "URW Gothic", "Symbola"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Calibri"), "Carlito");
			assert.strictEqual(oApp.GetSubstituteName("Cambria"), "Caladea");
			assert.strictEqual(oApp.GetSubstituteName("Arial"), "Liberation Sans");
			assert.strictEqual(oApp.GetSubstituteName("Georgia"), "Gelasio");
			assert.strictEqual(oApp.GetSubstituteName("Century Gothic"), "URW Gothic");
		});
		withInstalled(["Comic Neue"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Comic Sans MS"), "Comic Neue", "first installed substitute in order");
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

	QUnit.test("known but missing font without substitute keeps the dictionary matching", function (assert)
	{
		withInstalled(["Carlito", "Liberation Sans"], function ()
		{
			assert.strictEqual(oApp.GetSubstituteName("Tahoma"), "Tahoma", "selection by panose/metrics as before");
			assert.strictEqual(oApp.GetSubstituteName("Book Antiqua"), "Book Antiqua", "listed substitute not installed");
		});
	});
});
