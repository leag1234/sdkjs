/*
 * (c) Copyright Ascensio System SIA 2010-2026
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

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const eps = 0.01;
	const indent = 7.62;

	function createCalendarLikeTable(layout, noWrap, grid)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(215.9, 279.4);
		sectPr.SetPageMargins(25.4, 25.4, 25.4, 25.4);
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word12);

		// day column, spacer column, day column (the grid of a calendar: 864, 222, 864 twips)
		grid = grid || [15.24, 3.92, 15.24];
		let table = AscTest.CreateTable(2, 3, grid);
		if (undefined !== layout)
		{
			table.SetTableLayout(layout);
			table.Set_TableW(tblwidth_Mm, grid[0] + grid[1] + grid[2]);
		}
		logicDocument.PushToContent(table);

		for (let row = 0; row < 2; ++row)
		{
			for (let col = 0; col < 3; ++col)
			{
				let cell = table.GetRow(row).GetCell(col);
				cell.SetW(new CTableMeasurement(tblwidth_Mm, grid[col]));
				if (noWrap)
					cell.SetNoWrap(true);
				if (1 === col)
					continue;

				let p = cell.GetContent().GetElement(0);
				p.SetParagraphIndent({FirstLine : indent});
				AscTest.AddTextToParagraph(p, 0 === row ? "Sun" : "11");
			}
		}
		AscTest.Recalculate();
		return table;
	}

	function getLines(table, row, col)
	{
		return table.GetRow(row).GetCell(col).GetContent().GetElement(0).Lines.length;
	}

	QUnit.module("Autofit table with a first line indent");

	QUnit.test("Test: columns grow to hold a word next to the first line indent", function (assert)
	{
		let table = createCalendarLikeTable(undefined, false);
		let word  = 2 * 5 + indent + 2 * 1.9;
		assert.ok(table.TableGridCalc[0] >= word - eps, "Column 0 holds the word and the indent");
		assert.ok(table.TableGridCalc[2] >= word - eps, "Column 2 holds the word and the indent");
		assert.ok(table.TableGridCalc[0] >= 15.24 - eps, "Column 0 is not narrower than its grid width");
		assert.close(table.TableGridCalc[1], 3.92, eps, "Empty spacer column keeps its width");
		assert.strictEqual(getLines(table, 1, 0), 1, "Number is on a single line");
	});

	QUnit.test("Test: no-wrap cells keep the word on one line", function (assert)
	{
		let table = createCalendarLikeTable(undefined, true);
		assert.ok(table.TableGridCalc[0] >= 3 * 5 + indent + 2 * 1.9 - eps, "Column 0 holds the unwrapped text");
		assert.strictEqual(getLines(table, 0, 0), 1, "Day name is on a single line");
		assert.strictEqual(getLines(table, 1, 0), 1, "Number is on a single line");
	});

	QUnit.test("Test: columns that already hold the text keep their grid width", function (assert)
	{
		let table = createCalendarLikeTable(undefined, false, [40, 3.92, 40]);
		assert.close(table.TableGridCalc[0], 40, eps, "Column 0 keeps its grid width");
		assert.close(table.TableGridCalc[1], 3.92, eps, "Column 1 keeps its grid width");
		assert.close(table.TableGridCalc[2], 40, eps, "Column 2 keeps its grid width");
	});

	QUnit.test("Test: fixed layout ignores the first line indent", function (assert)
	{
		let table = createCalendarLikeTable(tbllayout_Fixed, false);
		let fixedWithIndent = table.TableGridCalc.slice();

		// the same table without the indent
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(215.9, 279.4);
		sectPr.SetPageMargins(25.4, 25.4, 25.4, 25.4);
		let grid   = [15.24, 3.92, 15.24];
		let plain  = AscTest.CreateTable(2, 3, grid);
		plain.SetTableLayout(tbllayout_Fixed);
		plain.Set_TableW(tblwidth_Mm, grid[0] + grid[1] + grid[2]);
		logicDocument.PushToContent(plain);
		for (let row = 0; row < 2; ++row)
		{
			for (let col = 0; col < 3; ++col)
			{
				let cell = plain.GetRow(row).GetCell(col);
				cell.SetW(new CTableMeasurement(tblwidth_Mm, grid[col]));
				if (1 !== col)
					AscTest.AddTextToParagraph(cell.GetContent().GetElement(0), 0 === row ? "Sun" : "11");
			}
		}
		AscTest.Recalculate();

		for (let col = 0; col < 3; ++col)
			assert.close(fixedWithIndent[col], plain.TableGridCalc[col], eps, "Column " + col + " does not depend on the indent in a fixed table");
	});
});
