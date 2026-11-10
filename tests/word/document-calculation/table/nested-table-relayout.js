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
	const eps = 0.001;
	const tw = 25.4 / 1440;

	function createTables()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(215.9, 279.4);
		sectPr.SetPageMargins(25.4, 25.4, 25.4, 25.4);

		let outer = AscTest.CreateTable(1, 2, [3488 * tw, 3215 * tw]);
		outer.Set_TableW(tblwidth_Pct, 70);
		outer.Set_TableAlign(align_Center);
		logicDocument.PushToContent(outer);
		let cell1 = outer.GetRow(0).GetCell(0);
		let cell2 = outer.GetRow(0).GetCell(1);
		cell1.SetW(new CTableMeasurement(tblwidth_Mm, 4788 * tw));
		cell2.SetW(new CTableMeasurement(tblwidth_Mm, 4788 * tw));

		let inner = AscTest.CreateTable(2, 2, [1343 * tw, 1267 * tw]);
		inner.Set_TableW(tblwidth_Pct, 80);
		inner.Set_TableAlign(align_Center);
		cell1.GetContent().AddToContent(0, inner);
		let ic = [[inner.GetRow(0).GetCell(0), inner.GetRow(0).GetCell(1)], [inner.GetRow(1).GetCell(0), inner.GetRow(1).GetCell(1)]];
		for (let r = 0; r < 2; ++r)
		{
			ic[r][0].SetW(new CTableMeasurement(tblwidth_Mm, 1678 * tw));
			ic[r][1].SetW(new CTableMeasurement(tblwidth_Mm, 1629 * tw));
		}
		function setText(cell, texts)
		{
			let content = cell.GetContent();
			content.ClearContent(true);
			for (let i = 0; i < texts.length; ++i)
			{
				let p = AscTest.CreateParagraph();
				content.AddToContent(i, p);
				if (texts[i])
					AscTest.AddTextToParagraph(p, texts[i]);
			}
		}
		setText(ic[0][0], ["One", "Three"]);
		setText(ic[0][1], ["Two"]);
		setText(ic[1][0], [""]);
		setText(ic[1][1], ["Four"]);
		setText(cell2, ["To the left is a table inside a table, with some cells merged."]);
		ic[0][0].SetVMerge(vmerge_Restart);
		ic[1][0].SetVMerge(vmerge_Continue);
		return {outer : outer, inner : inner};
	}

	function relayoutGrids(t, count)
	{
		let result = [];
		for (let i = 0; i < count; ++i)
		{
			AscTest.Recalculate();
			result.push({outer : t.outer.TableGridCalc.slice(), inner : t.inner.TableGridCalc.slice()});
		}
		return result;
	}

	QUnit.module("Nested table relayout");

	QUnit.test("Issue 66: nested percent table keeps its grid across relayouts", function (assert)
	{
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word12);
		let t = createTables();
		let grids = relayoutGrids(t, 8);
		grids.forEach(function (g) { console.log("DBG relayout " + JSON.stringify(g)); });
		for (let i = 1; i < grids.length; ++i)
		{
			for (let col = 0; col < 2; ++col)
			{
				assert.close(grids[i].outer[col], grids[0].outer[col], eps, "outer column " + col + " is stable on relayout " + i);
				assert.close(grids[i].inner[col], grids[0].inner[col], eps, "inner column " + col + " is stable on relayout " + i);
			}
		}
		let last = grids[grids.length - 1];
		assert.strictEqual(last.inner.length, 2, "inner table keeps two columns");
		assert.ok(last.inner[1] > 5, "second inner column is not collapsed");
	});
});
