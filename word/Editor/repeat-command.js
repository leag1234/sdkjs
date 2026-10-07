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

/**
 * Repeat (F4, and Ctrl+Y when there is nothing to redo), as in Microsoft Word: the last
 * repeatable command of the user is applied again to the current selection or cursor.
 *
 * The command is remembered as a closure that calls the same entry point again with the same
 * arguments, so a repeat goes through the same lock check, history action (one undo step)
 * and revision tracking as the original command. Recorded entry points:
 *  - asc_docs_api methods of the ribbon, the mini toolbar, the context menus and the dialogs
 *    (REPEATABLE_API_METHODS), recorded from the public (quoted) API names;
 *  - keyboard shortcuts (executeShortcut) that change the document;
 *  - typing (EnterText: the text typed since the last other command), Backspace, Delete, Enter.
 * Navigation, selection, undo/redo, save, zoom and view commands are not recorded.
 */
(function(window)
{
	const Type = Asc.c_oAscDocumentShortcutType;

	// Commands of the public API that the user can repeat. Each is replayed with its arguments.
	const REPEATABLE_API_METHODS = [
		// characters
		"put_TextPrBold", "put_TextPrItalic", "put_TextPrUnderline", "put_TextPrStrikeout", "put_TextPrDStrikeout",
		"put_TextPrFontName", "put_TextPrFontSize", "put_TextPrBaseline", "put_TextPrCaps", "put_TextPrSmallCaps",
		"put_TextPrSpacing", "put_TextPrPosition", "put_TextPrLang", "put_TextColor", "put_LineHighLight", "SetMarkerFormat",
		"FontSizeIn", "FontSizeOut", "asc_ChangeTextCase", "ClearFormating", "asc_SetFontDialogProps",
		// paragraphs
		"put_PrAlign", "put_PrLineSpacing", "put_LineSpacingBeforeAfter", "asc_addSpaceBeforeParagraph",
		"asc_addSpaceAfterParagraph", "asc_removeSpaceBeforeParagraph", "asc_removeSpaceAfterParagraph",
		"put_PrIndent", "put_PrIndentRight", "put_PrFirstLineIndent", "IncreaseIndent", "DecreaseIndent",
		"paraApply", "put_Borders", "put_ParagraphShade", "put_Style", "put_ListType", "put_ListTypeCustom",
		"put_KeepLines", "put_KeepNext", "put_PageBreak", "put_WidowControl", "put_AddSpaceBetweenPrg",
		"put_ParagraphOutlineLvl", "asc_setRtlTextDirection", "asc_SetParagraphSuppressLineNumbers",
		// tables
		"put_Table", "tblApply", "addRowAbove", "addRowBelow", "addColumnLeft", "addColumnRight", "remRow",
		"remColumn", "MergeCells", "SplitCell",
		// insertions
		"put_AddPageBreak", "put_AddColumnBreak", "add_SectionBreak", "asc_insertSymbol", "add_Hyperlink",
		"asc_addComment", "AddImageUrl", "put_PageNum", "asc_AddMath", "AddTextArt", "asc_addDateTime", "asc_AddPageCount",
		// delete, cut, paste
		"asc_SelectionCut", "asc_PasteData"
	];

	// Arguments kept for the repeat: the pasted data is copied (the paste consumes the DOM element) and
	// the callbacks of the original paste are dropped
	// The highlight button applies the colour to the selection and switches the marker on: the repeat
	// applies the colour to the new selection without the marker
	const REPLAYED_AS = {
		SetMarkerFormat : function(args)
		{
			return args[0] ? {name : "put_LineHighLight", args : args.slice(1, 5)} : null;
		}
	};

	const ARGUMENT_FILTERS = {
		asc_PasteData : function(args)
		{
			let data = args[1];
			if (data && data.cloneNode)
				data = data.cloneNode(true);

			return [args[0], data, args[2], args[3], args[4]];
		}
	};

	/**
	 * State of the Repeat command of a document.
	 * @param {CDocument} logicDocument
	 * @constructor
	 */
	function CRepeatCommand(logicDocument)
	{
		this.Document   = logicDocument;
		this.Last       = null;   // {kind, replay: function}
		this.Depth      = 0;      // > 0 while a command that is recorded as a whole is running
		this.Replaying  = false;
		this.TypingOpen = false;  // the last command is a typed run that further typing extends
	}
	CRepeatCommand.prototype.canRecord = function()
	{
		return !this.Replaying && 0 === this.Depth && !this.Document.IsViewMode() && !this.Document.IsViewModeInReview();
	};
	CRepeatCommand.prototype.canRepeat = function()
	{
		return !!this.Last;
	};
	CRepeatCommand.prototype.notify = function(canRepeat, kind)
	{
		let api = this.Document.Api;
		if (api && api.sync_CanRepeatCallback)
			api.sync_CanRepeatCallback(canRepeat, kind);
	};
	CRepeatCommand.prototype.set = function(command)
	{
		this.Last       = command;
		this.TypingOpen = false;
		this.notify(true, command.kind);
	};
	CRepeatCommand.prototype.run = function(fn)
	{
		this.Depth++;
		try
		{
			return fn();
		}
		finally
		{
			this.Depth--;
		}
	};
	/**
	 * Applies the last command again.
	 * @returns {boolean} true when there was a command to repeat
	 */
	CRepeatCommand.prototype.execute = function()
	{
		if (!this.Last || this.Replaying || this.Document.IsViewMode() || this.Document.IsViewModeInReview())
			return false;

		let last = this.Last;
		this.TypingOpen = false;
		this.Replaying  = true;
		try
		{
			last.replay();
		}
		finally
		{
			this.Replaying = false;
		}
		return true;
	};
	CRepeatCommand.prototype.closeTyping = function()
	{
		this.TypingOpen = false;
	};
	CRepeatCommand.prototype.onTyped = function(codePoints)
	{
		if (!codePoints || (Array.isArray(codePoints) && !codePoints.length))
			return;

		let added = Array.isArray(codePoints) ? codePoints.slice() : (typeof codePoints === "string" ? codePoints.codePointsArray() : [codePoints]);
		if (this.TypingOpen && this.Last && "typing" === this.Last.kind)
		{
			this.Last.codePoints = this.Last.codePoints.concat(added);
			return;
		}

		let doc    = this.Document;
		let record = {
			kind       : "typing",
			codePoints : added,
			replay     : function()
			{
				doc.EnterText(record.codePoints.slice());
			}
		};
		this.set(record);
		this.TypingOpen = true;
	};
	/**
	 * Backspace right after typing removes the last typed character from the typed run, as in Word
	 * @returns {boolean} true when the key was absorbed by the typed run
	 */
	CRepeatCommand.prototype.onBackspaceInTyping = function()
	{
		if (!this.TypingOpen || !this.Last || "typing" !== this.Last.kind || this.Document.IsSelectionUse())
			return false;

		this.Last.codePoints.pop();
		if (!this.Last.codePoints.length)
		{
			this.Last       = null;
			this.TypingOpen = false;
			this.notify(false);
		}
		return true;
	};
	/**
	 * A key of the document (Backspace, Delete, Enter) is replayed as a key event
	 */
	CRepeatCommand.prototype.setKey = function(e)
	{
		let doc   = this.Document;
		let copy  = Object.assign(Object.create(Object.getPrototypeOf(e)), e);
		let kind  = 8 === e.KeyCode ? "backspace" : (46 === e.KeyCode ? "delete" : "enter");
		this.set({
			kind   : kind,
			replay : function()
			{
				doc.OnKeyDown(Object.assign(Object.create(Object.getPrototypeOf(copy)), copy));
			}
		});
	};

	CDocument.prototype.GetRepeatCommand = function()
	{
		if (!this.RepeatCommand)
			this.RepeatCommand = new CRepeatCommand(this);
		return this.RepeatCommand;
	};

	function getHistorySignature(doc)
	{
		let history = doc.History;
		let point   = history.Points[history.Index];
		return history.Index + ":" + (point && point.Items ? point.Items.length : 0);
	}

	// Typed text, deleting and Enter, closing of the typed run by the other keys and the mouse
	const originalOnKeyDown = CDocument.prototype.OnKeyDown;
	CDocument.prototype.OnKeyDown = function(e)
	{
		let repeat = this.GetRepeatCommand();
		if (repeat.Replaying || repeat.Depth)
			return originalOnKeyDown.apply(this, arguments);

		let code = e.KeyCode;
		let ctrl = e.CtrlKey || e.MacCmdKey || e.AltKey;
		let isModifier = 16 === code || 17 === code || 18 === code || 20 === code || 91 === code || 93 === code || 224 === code;
		let isTyping   = !ctrl || (e.CtrlKey && e.AltKey && !e.MacCmdKey && e.AltGr);
		isTyping = isTyping && (32 === code || (code >= 48 && code <= 90) || (code >= 96 && code <= 111) || (code >= 186 && code <= 222) || 229 === code);

		let isEdit = 8 === code || (46 === code && !e.ShiftKey) || (13 === code && !e.ShiftKey && !ctrl);
		if (8 === code && repeat.onBackspaceInTyping())
		{
			// the typed run stays open (and shrinks): Backspace is not a command of its own
			isEdit = false;
			isModifier = true;
		}

		if (!isModifier && !isTyping)
			repeat.closeTyping();

		if (!isEdit || this.IsViewMode() || this.IsViewModeInReview())
			return originalOnKeyDown.apply(this, arguments);

		let before = getHistorySignature(this);
		let result = originalOnKeyDown.apply(this, arguments);
		if (getHistorySignature(this) !== before)
			repeat.setKey(e);

		return result;
	};

	const originalOnMouseDown = CDocument.prototype.OnMouseDown;
	CDocument.prototype.OnMouseDown = function()
	{
		this.GetRepeatCommand().closeTyping();
		return originalOnMouseDown.apply(this, arguments);
	};

	const originalEnterText = CDocument.prototype.EnterText;
	CDocument.prototype.EnterText = function(value)
	{
		let repeat = this.GetRepeatCommand();
		let record = repeat.canRecord();
		let result = originalEnterText.apply(this, arguments);
		if (result && record)
			repeat.onTyped(value);

		return result;
	};

	// Replacements made by auto-correct while typing are part of the typed text, not new commands
	const originalCorrectEnterText = CDocument.prototype.CorrectEnterText;
	CDocument.prototype.CorrectEnterText = function()
	{
		let args = arguments;
		return this.GetRepeatCommand().run(() => originalCorrectEnterText.apply(this, args));
	};

	["Document_Undo", "Document_Redo"].forEach(function(name)
	{
		let original = CDocument.prototype[name];
		CDocument.prototype[name] = function()
		{
			this.GetRepeatCommand().closeTyping();
			return original.apply(this, arguments);
		};
	});

	// Shortcuts: toggles are recorded with the value they applied, so that the repeat applies the
	// formatting instead of toggling it on a paragraph that already has it (Word)
	const SHORTCUT_TOGGLES = {};
	function apiCall(name)
	{
		let args = Array.prototype.slice.call(arguments, 1);
		return function(doc)
		{
			return {kind : name, replay : function() { doc.Api[name].apply(doc.Api, args); }};
		};
	}
	function textPrToggle(property, apiName, getValue)
	{
		return function(doc)
		{
			let textPr = doc.GetCalculatedTextPr();
			if (!textPr)
				return null;

			let value = getValue(textPr);
			return function() { return apiCall(apiName, value)(doc); };
		};
	}
	SHORTCUT_TOGGLES[Type.Bold]        = textPrToggle("Bold", "put_TextPrBold", tp => tp.Bold !== true);
	SHORTCUT_TOGGLES[Type.Italic]      = textPrToggle("Italic", "put_TextPrItalic", tp => tp.Italic !== true);
	SHORTCUT_TOGGLES[Type.Underline]   = textPrToggle("Underline", "put_TextPrUnderline", tp => tp.Underline !== true);
	SHORTCUT_TOGGLES[Type.Strikeout]   = textPrToggle("Strikeout", "put_TextPrStrikeout", tp => tp.Strikeout !== true);
	SHORTCUT_TOGGLES[Type.Superscript] = textPrToggle("VertAlign", "put_TextPrBaseline", tp => tp.VertAlign === AscCommon.vertalign_SuperScript ? AscCommon.vertalign_Baseline : AscCommon.vertalign_SuperScript);
	SHORTCUT_TOGGLES[Type.Subscript]   = textPrToggle("VertAlign", "put_TextPrBaseline", tp => tp.VertAlign === AscCommon.vertalign_SubScript ? AscCommon.vertalign_Baseline : AscCommon.vertalign_SubScript);
	// (all caps and small caps, ResetChar... are replayed as shortcuts: they do not depend on the target)
	function paraAlign(doc)
	{
		return function()
		{
			let paraPr = doc.GetCalculatedParaPr();
			return paraPr && undefined !== paraPr.Jc ? apiCall("put_PrAlign", paraPr.Jc)(doc) : null;
		};
	}
	SHORTCUT_TOGGLES[Type.CenterPara]  = paraAlign;
	SHORTCUT_TOGGLES[Type.LeftPara]    = paraAlign;
	SHORTCUT_TOGGLES[Type.RightPara]   = paraAlign;
	SHORTCUT_TOGGLES[Type.JustifyPara] = paraAlign;
	function styleShortcut(styleName)
	{
		return function(doc) { return function() { return apiCall("put_Style", styleName)(doc); }; };
	}
	SHORTCUT_TOGGLES[Type.ApplyHeading1] = styleShortcut("Heading 1");
	SHORTCUT_TOGGLES[Type.ApplyHeading2] = styleShortcut("Heading 2");
	SHORTCUT_TOGGLES[Type.ApplyHeading3] = styleShortcut("Heading 3");

	// Shortcuts whose command does not depend on the target: the shortcut itself is repeated
	const SHORTCUT_REPLAYED = {};
	[
		Type.InsertPageBreak, Type.InsertColumnBreak, Type.InsertLineBreak, Type.InsertTableBreak,
		Type.NonBreakingSpace, Type.NonBreakingHyphen, Type.OptionalHyphen, Type.EmDash, Type.EnDash,
		Type.HorizontalEllipsis, Type.CopyrightSign, Type.RegisteredSign, Type.TrademarkSign, Type.EuroSign,
		Type.InsertInvertedQuestionMark, Type.InsertInvertedExclamationMark,
		Type.InsertPageNumber, Type.InsertPageField, Type.InsertDateField, Type.InsertTimeField,
		Type.ResetChar, Type.ResetParagraph, Type.ApplyListBullet, Type.Indent, Type.UnIndent,
		Type.IncreaseFontSize, Type.DecreaseFontSize, Type.IncreaseFontSizeByPoint, Type.DecreaseFontSizeByPoint,
		Type.AllCaps, Type.SmallCaps, Type.ChangeCase,
		Type.HangingIndent, Type.RemoveHangingIndent,
		Type.LineSpacingSingle, Type.LineSpacingDouble, Type.LineSpacingOneAndHalf,
		Type.ApplyNormalStyle, Type.PromoteHeading, Type.DemoteHeading,
		Type.ParagraphRtl, Type.ParagraphLtr, Type.InsertTableRowAbove, Type.PasteFormat,
		Type.SymbolFont
	].forEach(function(type) { if (undefined !== type) SHORTCUT_REPLAYED[type] = true; });

	const originalExecuteShortcut = CDocument.prototype.executeShortcut;
	CDocument.prototype.executeShortcut = function(type)
	{
		let repeat = this.GetRepeatCommand();

		// Ctrl+Y (Cmd+Y): redo, and repeat when there is nothing to redo
		if (Type.EditRedo === type && !this.History.Can_Redo() && repeat.canRepeat() && !this.IsViewMode() && !this.IsViewModeInReview())
		{
			repeat.execute();
			return true;
		}
		if (Type.RepeatCommand === type)
		{
			if (!this.IsViewMode() && !this.IsViewModeInReview())
				repeat.execute();

			return true;
		}

		let record = repeat.canRecord();
		let toggle = record ? SHORTCUT_TOGGLES[type] : null;
		let post   = toggle ? toggle(this) : null;
		let doc    = this;
		let result = repeat.run(() => originalExecuteShortcut.apply(this, arguments));
		if (result && record)
		{
			if (post)
			{
				let command = post();
				command && repeat.set(command);
			}
			else if (SHORTCUT_REPLAYED[type])
			{
				repeat.set({
					kind   : "shortcut",
					replay : function() { originalExecuteShortcut.call(doc, type); }
				});
			}
		}
		return result;
	};

	/**
	 * Wraps the repeatable methods of the public API (called with the quoted names by the interface)
	 * @param {Function} apiClass asc_docs_api
	 */
	AscCommon.installRepeatCommand = function(apiClass)
	{
		const proto     = apiClass.prototype;
		const originals = {};
		REPEATABLE_API_METHODS.forEach(function(name)
		{
			if (proto[name])
				originals[name] = proto[name];
		});
		REPEATABLE_API_METHODS.forEach(function(name)
		{
			let original = originals[name];
			if (!original)
				return;

			proto[name] = function()
			{
				let doc = this.private_GetLogicDocument ? this.private_GetLogicDocument() : null;
				if (!doc || !doc.GetRepeatCommand)
					return original.apply(this, arguments);

				let repeat = doc.GetRepeatCommand();
				if (!repeat.canRecord())
					return original.apply(this, arguments);

				let callArgs = Array.prototype.slice.call(arguments);
				let replayAs = REPLAYED_AS[name] ? REPLAYED_AS[name](callArgs) : {name : name, args : callArgs};
				if (replayAs && ARGUMENT_FILTERS[replayAs.name])
					replayAs.args = ARGUMENT_FILTERS[replayAs.name](replayAs.args);

				let api    = this;
				let result = repeat.run(() => original.apply(api, callArgs));
				if (replayAs && repeat.canRecord())
				{
					repeat.set({
						kind   : replayAs.name,
						replay : function()
						{
							let args = ARGUMENT_FILTERS[replayAs.name] ? ARGUMENT_FILTERS[replayAs.name](replayAs.args) : replayAs.args;
							originals[replayAs.name].apply(api, args);
						}
					});
				}
				return result;
			};
		});
		proto["sync_CanRepeatCallback"] = proto.sync_CanRepeatCallback = function(canRepeat, kind)
		{
			this.sendEvent("asc_onCanRepeat", canRepeat, kind);
		};
		proto["asc_Repeat"] = proto.asc_Repeat = function()
		{
			let doc = this.private_GetLogicDocument();
			return !!(doc && doc.CanEdit() && doc.GetRepeatCommand().execute());
		};
		proto["asc_CanRepeat"] = proto.asc_CanRepeat = function()
		{
			let doc = this.private_GetLogicDocument();
			return !!(doc && doc.CanEdit() && doc.GetRepeatCommand().canRepeat());
		};
	};

	AscCommon.installRepeatCommand(AscCommon.DocumentEditorApi);
})(window);
