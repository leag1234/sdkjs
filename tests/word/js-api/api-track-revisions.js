/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

$(function ()
{
	QUnit.module("Track revisions: local/global flags in the builder API");

	let editor = AscTest.Editor;

	function setupEditor()
	{
		// The test editor is a baseEditorsApi: borrow the Word API implementations of the flag accessors
		editor.asc_GetLocalTrackRevisions = AscCommon.DocumentEditorApi.prototype.asc_GetLocalTrackRevisions.bind(editor);
		editor.asc_SetLocalTrackRevisions = AscCommon.DocumentEditorApi.prototype.asc_SetLocalTrackRevisions.bind(editor);
		editor.sync_OnTrackRevisionsChange = function(){};
		AscBuilder.Word.init();

		let userInfo = new Asc.asc_CUserInfo();
		userInfo.put_Id("uid-1");
		userInfo.put_FullName("John Smith");
		let docInfo = new Asc.asc_CDocInfo();
		docInfo.put_UserInfo(userInfo);
		editor.DocInfo = docInfo;
	}

	function getParagraphRunReviewType(apiParagraph, runIndex)
	{
		return apiParagraph.GetElement(runIndex).Run.GetReviewType();
	}

	QUnit.test("Api.asc_GetLocalTrackRevisions reads the local flag without changing it", function (assert)
	{
		setupEditor();
		let Api = AscBuilder.Word.Api;
		let logicDocument = AscTest.GetLogicDocument();

		assert.notStrictEqual(Api.asc_GetLocalTrackRevisions, Api.asc_SetLocalTrackRevisions, "getter and setter are different functions");

		logicDocument.SetLocalTrackRevisions(true);
		assert.strictEqual(Api.asc_GetLocalTrackRevisions(), true, "getter returns the local flag");
		assert.strictEqual(logicDocument.GetLocalTrackRevisions(), true, "getter does not reset the local flag");
		assert.strictEqual(logicDocument.IsTrackRevisions(), true, "tracking is still on after reading the flag");

		Api.asc_SetLocalTrackRevisions(false);
		assert.strictEqual(logicDocument.GetLocalTrackRevisions(), false, "setter changes the local flag");
		assert.strictEqual(Api.asc_GetLocalTrackRevisions(), false, "getter returns the new local flag");

		logicDocument.SetLocalTrackRevisions(null);
	});

	QUnit.test("SetAssistantTrackRevisions tracks changes even when tracking is switched off locally", function (assert)
	{
		setupEditor();
		let logicDocument = AscTest.GetLogicDocument();
		let doc = AscTest.JsApi.GetDocument();

		logicDocument.SetGlobalTrackRevisions(false);
		logicDocument.SetLocalTrackRevisions(false);

		doc.SetAssistantTrackRevisions(true, "Elixir AI");
		assert.strictEqual(doc.IsTrackRevisions(), true, "tracking is effective while the assistant works");

		let p = AscTest.JsApi.CreateParagraph();
		p.AddText("inserted by the assistant");
		doc.Push(p);
		assert.strictEqual(getParagraphRunReviewType(p, 0), reviewtype_Add, "new text is a tracked insertion");
		assert.strictEqual(p.GetElement(0).Run.GetReviewInfo().GetUserName(), "Elixir AI", "the revision author is the assistant");

		doc.SetAssistantTrackRevisions(false);
		assert.strictEqual(logicDocument.GetLocalTrackRevisions(), false, "local flag restored");
		assert.strictEqual(logicDocument.GetGlobalTrackRevisions(), false, "global flag restored");
		assert.strictEqual(editor.DocInfo.get_UserInfo().get_FullName(), "John Smith", "user name restored");
		assert.strictEqual(editor.DocInfo.get_UserInfo().get_Id(), "uid-1", "user id restored");

		logicDocument.SetLocalTrackRevisions(null);
	});

	QUnit.test("SetAssistantTrackRevisions restores a local-only tracking session exactly", function (assert)
	{
		setupEditor();
		let logicDocument = AscTest.GetLogicDocument();
		let doc = AscTest.JsApi.GetDocument();

		// "Track changes for me only": local flag on, global flag off
		logicDocument.SetGlobalTrackRevisions(false);
		logicDocument.SetLocalTrackRevisions(true);

		doc.SetAssistantTrackRevisions(true, "Elixir AI");
		assert.strictEqual(doc.IsTrackRevisions(), true, "tracking is on");
		doc.SetAssistantTrackRevisions(false);

		assert.strictEqual(logicDocument.GetGlobalTrackRevisions(), false, "tracking is not switched on for everyone");
		assert.strictEqual(logicDocument.GetLocalTrackRevisions(), true, "local flag restored");

		// No local flag, global flag on
		logicDocument.SetLocalTrackRevisions(null);
		logicDocument.SetGlobalTrackRevisions(true);
		doc.SetAssistantTrackRevisions(true, "Elixir AI");
		doc.SetAssistantTrackRevisions(false);
		assert.strictEqual(logicDocument.GetLocalTrackRevisions(), null, "local flag stays unset");
		assert.strictEqual(logicDocument.GetGlobalTrackRevisions(), true, "global flag restored");

		logicDocument.SetGlobalTrackRevisions(false);
	});
});
