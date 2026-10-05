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

// CDocument.CheckLastParagraphAfterLoad: a body read from a file that ends with a table gets the empty final
// paragraph Word adds (it takes room in the layout, as in Word).
$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();

	QUnit.module("Document end after load");

	QUnit.test("a final table is followed by an empty paragraph", function (assert)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);

		let table = AscTest.CreateTable(1, 1);
		table.GetRow(0).SetHeight(295, Asc.linerule_AtLeast);
		logicDocument.ReplaceContent([table]);
		assert.strictEqual(logicDocument.Content[logicDocument.Content.length - 1], table, "ReplaceContent keeps the table last (as read)");

		logicDocument.CheckLastParagraphAfterLoad();
		assert.strictEqual(logicDocument.Content.length, 2, "one element added");
		assert.true(logicDocument.Content[1].IsParagraph(), "the last element is a paragraph");

		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "a table filling the page pushes the final paragraph to page 2 (Word)");

		logicDocument.CheckLastParagraphAfterLoad();
		assert.strictEqual(logicDocument.Content.length, 2, "nothing added when the body already ends with a paragraph");
	});
});
