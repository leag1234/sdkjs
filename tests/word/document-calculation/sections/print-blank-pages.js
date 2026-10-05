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

// Empty pages printed by Word before sections that must start on an odd/even page number
// (CDocument.IsPrintBlankPageBefore, used by the PDF export).
$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();

	// Three sections, one paragraph (one page) each. types/starts: section break type and page number start
	// (-1 = continue numbering) of sections 2 and 3.
	function setup(types, starts, evenAndOdd)
	{
		AscTest.ClearDocument();
		let finalSectPr = AscTest.GetFinalSection();
		finalSectPr.SetPageSize(400, 400);
		finalSectPr.SetPageMargins(50, 50, 50, 50);

		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		let p3 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(p3);

		let s1 = new AscWord.SectPr(logicDocument);
		s1.Copy(finalSectPr);
		s1.SetPageNumStart(-1);
		let s2 = new AscWord.SectPr(logicDocument);
		s2.Copy(finalSectPr);
		s2.Set_Type(types[0]);
		s2.SetPageNumStart(starts[0]);
		finalSectPr.Set_Type(types[1]);
		finalSectPr.SetPageNumStart(starts[1]);
		p1.Set_SectionPr(s1);
		p2.Set_SectionPr(s2);

		logicDocument.Set_DocumentEvenAndOddHeaders(!!evenAndOdd);
		AscTest.Recalculate();
	}

	function blanks()
	{
		let result = [];
		for (let i = 0; i < logicDocument.GetPagesCount(); ++i)
			result.push(logicDocument.IsPrintBlankPageBefore(i));
		return result;
	}

	const NextPage = c_oAscSectionBreakType.NextPage;
	const OddPage  = c_oAscSectionBreakType.OddPage;
	const EvenPage = c_oAscSectionBreakType.EvenPage;

	QUnit.module("Print: blank pages before sections", {
		afterEach : function () { logicDocument.Set_DocumentEvenAndOddHeaders(false); }
	});

	QUnit.test("next page breaks never add pages without odd/even headers", function (assert)
	{
		setup([NextPage, NextPage], [1, 1], false);
		assert.strictEqual(logicDocument.GetPagesCount(), 3, "3 pages");
		assert.deepEqual(blanks(), [false, false, false], "numbering restarted at 1 twice: no blank page (Word)");
	});

	QUnit.test("odd/even headers: restarting at an odd number after an odd page", function (assert)
	{
		setup([NextPage, NextPage], [1, 1], true);
		assert.deepEqual(blanks(), [false, true, true], "1 | blank | i | blank | 1 (Word)");

		setup([NextPage, NextPage], [-1, -1], true);
		assert.deepEqual(blanks(), [false, false, false], "continuous numbering 1, 2, 3: no blank page");

		setup([NextPage, NextPage], [2, 1], true);
		assert.deepEqual(blanks(), [false, false, false], "1 | 2 | 1: parity alternates");
	});

	QUnit.test("odd page and even page section breaks", function (assert)
	{
		setup([OddPage, OddPage], [-1, -1], false);
		assert.deepEqual(blanks(), [false, true, true], "odd page after page 1, then after page 3");

		setup([EvenPage, EvenPage], [-1, -1], false);
		assert.deepEqual(blanks(), [false, false, true], "even page after page 1: none; after page 2: blank, then page 4");
	});
});
