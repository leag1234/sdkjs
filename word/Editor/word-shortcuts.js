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
 * Microsoft Word keyboard shortcuts that the editor did not implement:
 * the commands behind the new c_oAscDocumentShortcutType values (customizable,
 * see Shortcuts.js) and a few keyboard mechanics handled before the regular
 * key processing (extend-selection mode, accent prefixes, table navigation).
 */
(function(window)
{
	const Type = Asc.c_oAscDocumentShortcutType;

	// Word's font size list (Grow Font / Shrink Font use Api.FontSizeIn/Out)
	const DEFAULT_TAB = 12.7; // 0.5 inch, Word's default tab stop

	function isMac()
	{
		return AscCommon.AscBrowser.isMacOs;
	}

	function getDefaultTab()
	{
		return (AscCommonWord.Default_Tab_Stop > 0.001 ? AscCommonWord.Default_Tab_Stop : DEFAULT_TAB);
	}

	// Date and time formats used by Alt+Shift+D / Alt+Shift+T, by language
	const DATE_FORMATS = {
		1033 : ["M/d/yyyy", "h:mm am/pm"],
		2057 : ["dd/MM/yyyy", "HH:mm"],
		1036 : ["dd/MM/yyyy", "HH:mm"],
		1031 : ["dd.MM.yyyy", "HH:mm"],
		3082 : ["dd/MM/yyyy", "H:mm"],
		1034 : ["dd/MM/yyyy", "H:mm"],
		1040 : ["dd/MM/yyyy", "HH:mm"],
		1046 : ["dd/MM/yyyy", "HH:mm"],
		2070 : ["dd/MM/yyyy", "HH:mm"],
		1043 : ["d-M-yyyy", "HH:mm"],
		1049 : ["dd.MM.yyyy", "H:mm"],
		1045 : ["dd.MM.yyyy", "HH:mm"],
		1041 : ["yyyy/MM/dd", "H:mm"],
		2052 : ["yyyy/M/d", "H:mm"]
	};

	// Word's "Ctrl+<accent>, letter" composition (US keyboard key names)
	const ACCENTS = {
		grave      : {a : "à", e : "è", i : "ì", o : "ò", u : "ù", A : "À", E : "È", I : "Ì", O : "Ò", U : "Ù"},
		acute      : {a : "á", e : "é", i : "í", o : "ó", u : "ú", y : "ý", A : "Á", E : "É", I : "Í", O : "Ó", U : "Ú", Y : "Ý", d : "ð", D : "Ð"},
		circumflex : {a : "â", e : "ê", i : "î", o : "ô", u : "û", A : "Â", E : "Ê", I : "Î", O : "Ô", U : "Û"},
		tilde      : {a : "ã", n : "ñ", o : "õ", A : "Ã", N : "Ñ", O : "Õ"},
		umlaut     : {a : "ä", e : "ë", i : "ï", o : "ö", u : "ü", y : "ÿ", A : "Ä", E : "Ë", I : "Ï", O : "Ö", U : "Ü", Y : "Ÿ"},
		ring       : {a : "å", A : "Å"},
		ligature   : {a : "æ", A : "Æ", o : "œ", O : "Œ", s : "ß"},
		cedilla    : {c : "ç", C : "Ç"},
		stroke     : {o : "ø", O : "Ø"}
	};

	function getAccentPrefix(e)
	{
		// Word for Mac uses the Option dead keys of the keyboard layout instead
		if (isMac() || !e.CtrlKey || e.AltKey || e.AltGr)
			return null;

		let code = e.KeyCode;
		if (!e.ShiftKey)
		{
			if (192 === code)
				return "grave";
			if (222 === code)
				return "acute";
			if (188 === code)
				return "cedilla";
			if (191 === code)
				return "stroke";
		}
		else
		{
			if (54 === code)
				return "circumflex";
			if (192 === code)
				return "tilde";
			if (186 === code || 59 === code)
				return "umlaut";
			if (50 === code)
				return "ring";
			if (55 === code)
				return "ligature";
		}
		return null;
	}

	function getCurrentTableCell(logicDocument)
	{
		let paragraph = logicDocument.GetCurrentParagraph();
		if (!paragraph || !paragraph.Parent || !paragraph.Parent.IsTableCellContent || !paragraph.Parent.IsTableCellContent())
			return null;

		return paragraph.Parent.GetParent();
	}

	function moveToCell(logicDocument, cell)
	{
		if (!cell)
			return;

		logicDocument.RemoveSelection();
		let paragraph = cell.GetContent().GetElement(0);
		if (paragraph && paragraph.IsParagraph())
		{
			paragraph.MoveCursorToStartPos(false);
			paragraph.Document_SetThisElementCurrent(true);
		}
		else
		{
			cell.GetContent().MoveCursorToStartPos(false);
		}
		logicDocument.UpdateSelection();
		logicDocument.UpdateInterface();
	}

	/**
	 * Keyboard mechanics handled before the regular processing of OnKeyDown.
	 * @returns {number|undefined} a keydownresult value when the key was handled
	 */
	CDocument.prototype.handleWordKeyDown = function(e)
	{
		const PreventAll = 0xFFFF;
		let keyCode = e.KeyCode;

		// Alt+Shift+F10 belongs to the interface (paste options), not to the Shift+F10 context menu
		if (121 === keyCode && e.ShiftKey && e.AltKey)
			return 0;

		// Alt+X (Windows, Linux; Cmd+Option+X on macOS): Unicode code <-> character
		if (88 === keyCode && e.AltKey && !e.ShiftKey && (isMac() ? e.MacCmdKey : (!e.CtrlKey && !e.MacCmdKey)) && this.private_ToggleUnicode())
		{
			this.UpdateSelection();
			return PreventAll;
		}

		// Ctrl+', Ctrl+`, Ctrl+Shift+: ... then a letter
		let accent = getAccentPrefix(e);
		if (accent)
		{
			this.WordPendingAccent = accent;
			return PreventAll;
		}
		if (this.WordPendingAccent && (e.CtrlKey || e.AltKey || 8 === keyCode || 9 === keyCode || 13 === keyCode || 27 === keyCode || (keyCode >= 33 && keyCode <= 40) || 46 === keyCode))
			this.WordPendingAccent = null;

		// F8 extend-selection mode: navigation keys extend the selection
		if (this.WordExtendSelection)
		{
			if (27 === keyCode)
			{
				this.WordExtendSelection = 0;
				return PreventAll;
			}
			if ((keyCode >= 33 && keyCode <= 40) && !e.ShiftKey)
				e.ShiftKey = true;
			else if (119 !== keyCode && 16 !== keyCode && 17 !== keyCode && 18 !== keyCode && 91 !== keyCode)
				this.WordExtendSelection = 0;
			// (typing, deleting or any other command leaves the mode, as in Word)
		}

		// Ctrl+Up: to the start of the current paragraph first (Word), then to the previous one
		if (38 === keyCode && !e.ShiftKey && !e.AltKey && (isMac() ? e.MacCmdKey : (e.CtrlKey && !e.MacCmdKey)))
		{
			let paragraph = this.GetCurrentParagraph();
			if (paragraph && !this.IsSelectionUse() && docpostype_DrawingObjects !== this.GetDocPosType() && !paragraph.IsCursorAtBegin())
			{
				paragraph.MoveCursorToStartPos(false);
				paragraph.Document_SetThisElementCurrent(true);
				this.UpdateSelection();
				return PreventAll;
			}
		}

		// Tables
		let cell = (keyCode >= 9 && keyCode <= 40) || 12 === keyCode ? getCurrentTableCell(this) : null;
		if (cell)
		{
			let table = cell.GetTable();
			let row   = cell.GetRow();
			let tableModifier = isMac() ? (e.MacCtrlKey && !e.MacCmdKey && !e.AltKey) : (e.AltKey && !e.CtrlKey);

			if (tableModifier && !e.ShiftKey && (35 === keyCode || 36 === keyCode))
			{
				// Alt+Home, Alt+End (Control+Home, Control+End on macOS): first, last cell of the row
				moveToCell(this, row.GetCell(36 === keyCode ? 0 : row.GetCellsCount() - 1));
				return PreventAll;
			}
			if (tableModifier && !e.ShiftKey && (33 === keyCode || 34 === keyCode))
			{
				// Alt+Page Up, Alt+Page Down: first, last cell of the column
				let cellIndex = cell.GetIndex();
				let rowIndex  = 33 === keyCode ? 0 : table.GetRowsCount() - 1;
				let target    = table.GetRow(rowIndex).GetCell(Math.min(cellIndex, table.GetRow(rowIndex).GetCellsCount() - 1));
				moveToCell(this, target);
				return PreventAll;
			}
			if (e.AltKey && e.ShiftKey && !e.CtrlKey && (35 === keyCode || 36 === keyCode))
			{
				// Shift+Alt+End, Shift+Alt+Home from the first or last cell: the row
				this.SelectTable(c_oAscTableSelectionType.Row);
				return PreventAll;
			}
			if (e.ShiftKey && !e.AltKey && !e.CtrlKey && (33 === keyCode || 34 === keyCode) && !this.IsSelectionUse())
			{
				// Shift+Page Down from the top cell, Shift+Page Up from the bottom cell: the column
				let rowIndex = row.GetIndex();
				if ((34 === keyCode && 0 === rowIndex) || (33 === keyCode && table.GetRowsCount() - 1 === rowIndex))
				{
					this.SelectTable(c_oAscTableSelectionType.Column);
					return PreventAll;
				}
			}
			if (12 === keyCode && e.AltKey)
			{
				// Alt+5 on the numeric keypad with Num Lock off (Clear): the whole table
				this.SelectTable(c_oAscTableSelectionType.Table);
				return PreventAll;
			}
			if (9 === keyCode && !e.ShiftKey && (isMac() ? (e.AltKey && !e.MacCmdKey) : (e.CtrlKey && !e.AltKey)))
			{
				// Ctrl+Tab (Option+Tab on macOS): a tab character in the cell
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
				{
					this.StartAction(AscDFH.historydescription_Document_AddTab);
					this.AddToParagraph(new AscWord.CRunTab());
					this.FinalizeAction();
				}
				return PreventAll;
			}
		}

		return undefined;
	};

	/**
	 * Applied to the next typed character after an accent prefix
	 * @returns {number[]} code points to insert
	 */
	CDocument.prototype.composeWordAccent = function(codePoints)
	{
		let accent = this.WordPendingAccent;
		this.WordPendingAccent = null;
		if (!accent || !codePoints || 1 !== codePoints.length)
			return codePoints;

		let composed = ACCENTS[accent][String.fromCodePoint(codePoints[0])];
		return composed ? [composed.codePointAt(0)] : codePoints;
	};

	CDocument.prototype.private_SetTextPrByWordShortcut = function(textPr)
	{
		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			return;

		this.StartAction(AscDFH.historydescription_Document_SetTextFontSize);
		this.AddToParagraph(new AscCommonWord.ParaTextPr(textPr));
		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.private_ChangeFontSizeByPoint = function(delta)
	{
		let textPr = this.GetCalculatedTextPr();
		if (!textPr || !textPr.FontSize)
			return;

		let size = Math.max(1, Math.min(1638, Math.round(textPr.FontSize) + delta));
		this.private_SetTextPrByWordShortcut({FontSize : size});
	};

	CDocument.prototype.private_CycleTextCase = function()
	{
		if (!this.IsSelectionUse())
			this.SelectCurrentWord();

		let text = this.GetSelectedText(false) || "";
		if (!text.trim())
			return;

		let upper = text.toLocaleUpperCase();
		let lower = text.toLocaleLowerCase();
		let caseType;
		if (text === upper && text !== lower)
			caseType = Asc.c_oAscChangeTextCaseType.CapitalizeWords;
		else if (text !== lower && /^(\P{L}*\p{Lu}\p{Ll}*)+\P{L}*$/u.test(text))
			caseType = Asc.c_oAscChangeTextCaseType.LowerCase;
		else
			caseType = Asc.c_oAscChangeTextCaseType.UpperCase;

		this.ChangeTextCase(caseType);
	};

	CDocument.prototype.private_SetHangingIndent = function(isIncrease)
	{
		let paraPr = this.GetCalculatedParaPr();
		if (!paraPr || !paraPr.Ind || this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			return;

		let tab       = getDefaultTab();
		let left      = paraPr.Ind.Left || 0;
		let firstLine = paraPr.Ind.FirstLine || 0;
		let firstAbs  = left + firstLine;

		let newLeft = isIncrease ? (Math.floor(left / tab + 1e-6) + 1) * tab : Math.max(0, (Math.ceil(left / tab - 1e-6) - 1) * tab);
		let newFirst = Math.min(0, firstAbs - newLeft);

		this.StartAction(AscDFH.historydescription_Document_SetParagraphIndent);
		this.SetParagraphIndent({Left : newLeft, FirstLine : newFirst});
		this.FinalizeAction();
	};

	CDocument.prototype.private_SetLineSpacingByWordShortcut = function(value)
	{
		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			return;

		this.StartAction(AscDFH.historydescription_Document_SetParagraphLineSpacing);
		this.SetParagraphSpacing({LineRule : Asc.linerule_Auto, Line : value});
		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.private_ApplyStyleByWordShortcut = function(styleId)
	{
		let style = styleId ? this.Styles.Get(styleId) : null;
		if (!style || this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			return;

		this.StartAction(AscDFH.historydescription_Document_SetParagraphStyle);
		this.SetParagraphStyle(style.GetName(), true);
		this.FinalizeAction();
	};

	CDocument.prototype.private_ChangeHeadingLevel = function(delta)
	{
		let paraPr = this.GetCalculatedParaPr();
		let styleId = paraPr ? paraPr.PStyle : null;
		let level = styleId ? this.Styles.GetHeadingLevelById(styleId) : -1;
		if (!(level >= 0))
			level = -1;

		let newLevel;
		if (-1 === level)
		{
			// body text: promote makes it Heading 1, demote does nothing
			if (delta > 0)
				return;
			newLevel = 0;
		}
		else
		{
			newLevel = level + delta;
			if (newLevel < 0)
				return;
		}

		if (newLevel > 8)
			this.private_ApplyStyleByWordShortcut(this.Styles.GetDefaultParagraph());
		else
			this.private_ApplyStyleByWordShortcut(this.Styles.GetDefaultHeading(newLevel));
	};

	CDocument.prototype.private_MoveParagraphsOrRow = function(isUp)
	{
		let cell = getCurrentTableCell(this);
		if (cell)
		{
			let table = cell.GetTable();
			let rowIndex = cell.GetRow().GetIndex();
			let target = isUp ? rowIndex - 1 : rowIndex + 1;
			if (target < 0 || target >= table.GetRowsCount())
				return;

			if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_Element_and_Type, Element : table, CheckType : AscCommon.changestype_Table_Properties}))
				return;

			this.StartAction(AscDFH.historydescription_Document_TableAddNewRowByTab);
			let moved = table.GetRow(isUp ? rowIndex - 1 : rowIndex + 1);
			let movedIndex = moved.GetIndex();
			table.private_RemoveRow(movedIndex);
			table.private_AddRow(isUp ? rowIndex : rowIndex, 0, true, moved);
			table.Internal_ReIndexing(0);
			this.Recalculate();
			this.UpdateSelection();
			this.UpdateInterface();
			this.FinalizeAction();
			return;
		}

		let paragraph = this.GetCurrentParagraph();
		if (!paragraph || docpostype_DrawingObjects === this.GetDocPosType())
			return;

		let content = paragraph.GetParent();
		if (!content || !content.Content)
			return;

		let start, end;
		if (content.IsSelectionUse() && content.Selection.StartPos !== undefined)
		{
			start = Math.min(content.Selection.StartPos, content.Selection.EndPos);
			end   = Math.max(content.Selection.StartPos, content.Selection.EndPos);
		}
		else
		{
			start = end = paragraph.GetIndex();
		}

		let neighbourIndex = isUp ? start - 1 : end + 1;
		if (neighbourIndex < 0 || neighbourIndex >= content.Content.length)
			return;

		let neighbour = content.Content[neighbourIndex];
		if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : [neighbour, paragraph], CheckType : AscCommon.changestype_Paragraph_Content}))
			return;

		this.StartAction(AscDFH.historydescription_Document_MoveParagraphByTab);
		content.Internal_Content_Remove(neighbourIndex, 1);
		if (isUp)
			content.Internal_Content_Add(end, neighbour);
		else
			content.Internal_Content_Add(start, neighbour);

		// (Internal_Content_Remove/Add keep CurPos and Selection on the moved paragraphs)

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.private_InsertFieldByWordShortcut = function(instruction)
	{
		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			return;

		let textPr = this.GetDirectTextPr().Copy();
		this.StartAction(AscDFH.historydescription_Document_AddDateTimeField);
		this.AddFieldWithInstruction(instruction, textPr);
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	};

	CDocument.prototype.private_GetDateTimeFormat = function(isTime)
	{
		let textPr = this.GetCalculatedTextPr();
		let lang = textPr && textPr.Lang && textPr.Lang.Val ? textPr.Lang.Val : 1033;
		let formats = DATE_FORMATS[lang] || DATE_FORMATS[1033];
		return formats[isTime ? 1 : 0];
	};

	CDocument.prototype.private_GetFieldsInReadingOrder = function()
	{
		let fields = this.GetAllFields(false).filter(function(field)
		{
			return field && field.IsValid && field.IsValid() && field.GetStartDocumentPosition;
		});
		fields.sort(function(a, b)
		{
			return AscWord.CompareDocumentPositions(a.GetStartDocumentPosition(), b.GetStartDocumentPosition());
		});
		return fields;
	};

	CDocument.prototype.private_GotoField = function(isNext)
	{
		let fields = this.private_GetFieldsInReadingOrder();
		if (!fields.length)
			return;

		let current = this.GetContentPosition(this.IsSelectionUse(), isNext ? false : true);
		let currentFields = this.GetCurrentComplexFields();
		let currentField = currentFields.length ? currentFields[currentFields.length - 1] : null;
		let target = null;
		if (isNext)
		{
			for (let i = 0; i < fields.length; ++i)
			{
				if (fields[i] !== currentField && AscWord.CompareDocumentPositions(fields[i].GetStartDocumentPosition(), current) >= 0)
				{
					target = fields[i];
					break;
				}
			}
		}
		else
		{
			for (let i = fields.length - 1; i >= 0; --i)
			{
				if (fields[i] !== currentField && AscWord.CompareDocumentPositions(fields[i].GetEndDocumentPosition(), current) <= 0)
				{
					target = fields[i];
					break;
				}
			}
		}

		if (!target)
			return;

		// select the whole field, as Word does
		this.RemoveSelection();
		this.SetSelectionByContentPositions(target.GetStartDocumentPosition(), target.GetEndDocumentPosition());
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.private_UnlinkFields = function()
	{
		let fields = this.IsSelectionUse() ? this.GetAllFields(true) : this.GetCurrentComplexFields();
		if (!fields || !fields.length || this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			return;

		this.StartAction(AscDFH.historydescription_Document_RemoveComplexFieldWrapper);
		for (let i = fields.length - 1; i >= 0; --i)
		{
			if (fields[i] && fields[i].RemoveFieldWrap)
				fields[i].RemoveFieldWrap();
		}
		this.RemoveSelection();
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.private_ToggleAllFieldCodes = function()
	{
		let fields = this.GetAllFields(false);
		if (!fields.length)
			return;

		let show = !fields.some(function(field) { return field.IsShowFieldCode && field.IsShowFieldCode(); });
		for (let i = 0; i < fields.length; ++i)
		{
			if (fields[i].IsShowFieldCode && fields[i].IsShowFieldCode() !== show)
				fields[i].BeginChar.SetShowFieldCode(show);
		}
		this.RecalculateFromStart(true);
		this.UpdateSelection();
	};

	// F8: extend mode on, then word, sentence, paragraph, document
	CDocument.prototype.private_ExtendSelectionStep = function()
	{
		if (!this.WordExtendSelection)
		{
			this.WordExtendSelection = 1;
			this.WordExtendStack = [];
			return;
		}

		let paragraph = this.GetCurrentParagraph();
		let level = this.WordExtendSelection;
		this.WordExtendStack.push(this.GetSelectionState());

		if (1 === level)
		{
			this.SelectCurrentWord();
			this.private_ExtendSelectionOverSpaces(paragraph);
			this.WordExtendSelection = 2;
		}
		else if (2 === level && paragraph)
		{
			this.private_SelectSentence(paragraph);
			this.WordExtendSelection = 3;
		}
		else if (3 === level && paragraph)
		{
			paragraph.SelectAll(1);
			paragraph.Document_SetThisElementCurrent(false);
			this.WordExtendSelection = 4;
		}
		else
		{
			this.SelectAll();
			this.WordExtendSelection = 5;
		}
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.private_ExtendSelectionOverSpaces = function(paragraph)
	{
		if (!paragraph || !this.IsSelectionUse())
			return;

		let text = paragraph.GetText({ParaEndToSpace : false});
		let end = this.private_GetOffsetInParagraph(paragraph, true);
		while (end < text.length && " " === text.charAt(end))
		{
			this.MoveCursorRight(true, false);
			++end;
		}
	};

	// Shift+F8: back to the previous F8 step
	CDocument.prototype.private_ShrinkSelectionStep = function()
	{
		if (!this.WordExtendStack || !this.WordExtendStack.length)
			return;

		this.SetSelectionState(this.WordExtendStack.pop());
		this.WordExtendSelection = Math.max(1, this.WordExtendSelection - 1);
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.private_SelectSentence = function(paragraph)
	{
		let text = paragraph.GetText({ParaEndToSpace : false}).replace(/[\r\n]+$/, "");
		let caret = this.private_GetOffsetInParagraph(paragraph);
		let start = 0, end = text.length;
		let re = /[.!?…]+["'”’)\]]*\s+/g, m;
		while ((m = re.exec(text)))
		{
			let boundary = m.index + m[0].length;
			if (boundary <= caret)
				start = boundary;
			else
			{
				end = boundary;
				break;
			}
		}
		this.private_SelectParagraphOffsets(paragraph, start, end);
	};

	CDocument.prototype.private_GetOffsetInParagraph = function(paragraph, isSelectionEnd)
	{
		let state = this.GetSelectionState();
		let target = paragraph.Get_ParaContentPos(this.IsSelectionUse(), !isSelectionEnd);
		paragraph.MoveCursorToStartPos(false);
		let offset = 0, guard = 10000;
		while (paragraph.Get_ParaContentPos(false, false).Compare(target) < 0 && guard-- > 0)
		{
			paragraph.MoveCursorRight(false, false);
			++offset;
		}
		this.SetSelectionState(state);
		return offset;
	};

	CDocument.prototype.private_SelectParagraphOffsets = function(paragraph, start, end)
	{
		paragraph.MoveCursorToStartPos(false);
		for (let i = 0; i < start; ++i)
			paragraph.MoveCursorRight(false, false);
		paragraph.Document_SetThisElementCurrent(false);
		this.RemoveSelection();
		paragraph.Document_SetThisElementCurrent(false);
		for (let i = start; i < end; ++i)
			this.MoveCursorRight(true, false);
	};

	// Ctrl+Alt+Page Up/Down, Ctrl+Alt+Shift+Page Down: first or last line of the visible area
	CDocument.prototype.private_MoveToScreenEdge = function(isBottom, addToSelect)
	{
		let drawingDocument = this.DrawingDocument;
		let wordControl = drawingDocument.m_oWordControl;
		let editorElement = wordControl && wordControl.m_oEditor ? wordControl.m_oEditor.HtmlElement : null;
		if (!editorElement || !editorElement.getBoundingClientRect)
			return;

		let rect = editorElement.getBoundingClientRect();
		let margin = 12; // a line partly hidden at the edge is not "on the screen"
		let pos = drawingDocument.ConvertCoordsFromCursor2(rect.left + 1, isBottom ? rect.bottom - margin : rect.top + margin);
		if (!pos || pos.Page < 0)
			return;

		if (docpostype_Content !== this.GetDocPosType())
		{
			this.RemoveSelection();
			this.SetDocPosType(docpostype_Content);
		}

		let scrollY = wordControl.m_dScrollY;
		this.CurPage = pos.Page;
		this.Controller.MoveCursorToXY(pos.X, pos.Y, pos.Page, addToSelect);
		if (isBottom)
			this.MoveCursorToEndOfLine(addToSelect);
		else
			this.MoveCursorToStartOfLine(addToSelect);
		if (wordControl.m_dScrollY !== scrollY && wordControl.m_oScrollVerApi)
			wordControl.m_oScrollVerApi.scrollToY(scrollY);
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.private_SelectFloatingObject = function()
	{
		let drawings = this.DrawingObjects.drawingObjects || [];
		let floating = drawings.filter(function(drawing)
		{
			return drawing && drawing.IsInline && !drawing.IsInline() && drawing.GraphicObj && !drawing.IsUseInDocument || (drawing && drawing.IsInline && !drawing.IsInline());
		});
		if (!floating.length)
			return;

		let current = this.CurPage || 0;
		floating.sort(function(a, b)
		{
			let pa = a.PageNum, pb = b.PageNum;
			let da = Math.abs(pa - current), db = Math.abs(pb - current);
			return da !== db ? da - db : (a.Y - b.Y);
		});
		this.SelectDrawings([floating[0]], this);
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.private_SelectNextSpellingError = function()
	{
		let current = this.GetCurrentParagraph();
		let paragraphs = this.GetAllParagraphs({All : true});
		let currentIndex = current ? paragraphs.indexOf(current) : -1;
		let currentPos = current ? current.Get_ParaContentPos(this.IsSelectionUse(), false) : null;
		let found = null;
		for (let i = Math.max(0, currentIndex); i < paragraphs.length && !found; ++i)
		{
			let paragraph = paragraphs[i];
			let spellChecker = paragraph.SpellChecker;
			if (!spellChecker)
				continue;

			for (let j = 0, n = spellChecker.GetElementsCount(); j < n; ++j)
			{
				let element = spellChecker.GetElement(j);
				if (!element || !element.IsWrong())
					continue;

				// in the current paragraph, only the errors after the caret (or after the selected error)
				if (i === currentIndex && currentPos && element.getStartParaPos().Compare(currentPos) < (this.IsSelectionUse() ? 1 : 0))
					continue;

				found = {paragraph : paragraph, element : element};
				break;
			}
		}
		if (!found)
			return false;

		this.RemoveSelection();
		found.paragraph.Selection.Use = true;
		found.paragraph.Set_SelectionContentPos(found.element.getStartParaPos(), found.element.getEndParaPos());
		found.paragraph.Document_SetThisElementCurrent(false);
		this.UpdateSelection();
		this.UpdateInterface();
		return true;
	};

	CDocument.prototype.private_OpenContextMenuAtTarget = function()
	{
		let pos = this.DrawingDocument.ConvertCoordsToCursorWR(this.TargetPos.X, this.TargetPos.Y, this.TargetPos.PageNum);
		this.Api.sync_ContextMenuCallback({Type : Asc.c_oAscContextMenuTypes.Common, X_abs : pos.X, Y_abs : pos.Y});
	};

	/**
	 * Commands of the Word shortcuts added to c_oAscDocumentShortcutType
	 * @returns {boolean} true when the shortcut was handled
	 */
	CDocument.prototype.executeWordShortcut = function(type)
	{
		switch (type)
		{
			case Type.IncreaseFontSizeByPoint:
				this.private_ChangeFontSizeByPoint(1);
				return true;
			case Type.DecreaseFontSizeByPoint:
				this.private_ChangeFontSizeByPoint(-1);
				return true;
			case Type.ChangeCase:
				this.private_CycleTextCase();
				return true;
			case Type.AllCaps:
			{
				let textPr = this.GetCalculatedTextPr();
				let caps = !(textPr && textPr.Caps);
				this.private_SetTextPrByWordShortcut(caps ? {Caps : true, SmallCaps : false} : {Caps : false});
				return true;
			}
			case Type.SmallCaps:
			{
				let textPr = this.GetCalculatedTextPr();
				let smallCaps = !(textPr && textPr.SmallCaps);
				this.private_SetTextPrByWordShortcut(smallCaps ? {SmallCaps : true, Caps : false} : {SmallCaps : false});
				return true;
			}
			case Type.SymbolFont:
				this.private_SetTextPrByWordShortcut({RFonts : {Ascii : {Name : "Symbol", Index : -1}, HAnsi : {Name : "Symbol", Index : -1}, CS : {Name : "Symbol", Index : -1}, EastAsia : {Name : "Symbol", Index : -1}}});
				return true;
			case Type.HangingIndent:
				this.private_SetHangingIndent(true);
				return true;
			case Type.RemoveHangingIndent:
				this.private_SetHangingIndent(false);
				return true;
			case Type.ResetParagraph:
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
				{
					this.StartAction(AscDFH.historydescription_Document_ClearFormatting);
					this.ClearParagraphFormatting(true, false);
					this.FinalizeAction();
				}
				return true;
			case Type.LineSpacingSingle:
				this.private_SetLineSpacingByWordShortcut(1);
				return true;
			case Type.LineSpacingDouble:
				this.private_SetLineSpacingByWordShortcut(2);
				return true;
			case Type.LineSpacingOneAndHalf:
				this.private_SetLineSpacingByWordShortcut(1.5);
				return true;
			case Type.ApplyNormalStyle:
				this.private_ApplyStyleByWordShortcut(this.Styles.GetDefaultParagraph());
				return true;
			case Type.PromoteHeading:
				this.private_ChangeHeadingLevel(-1);
				return true;
			case Type.DemoteHeading:
				this.private_ChangeHeadingLevel(1);
				return true;
			case Type.MoveParagraphUp:
				this.private_MoveParagraphsOrRow(true);
				return true;
			case Type.MoveParagraphDown:
				this.private_MoveParagraphsOrRow(false);
				return true;
			case Type.ToggleTrackRevisions:
				this.Api.asc_SetTrackRevisions(!this.IsTrackRevisions());
				return true;
			case Type.InsertDateField:
				this.private_InsertFieldByWordShortcut("DATE \\@ \"" + this.private_GetDateTimeFormat(false) + "\"");
				return true;
			case Type.InsertTimeField:
				this.private_InsertFieldByWordShortcut("TIME \\@ \"" + this.private_GetDateTimeFormat(true) + "\"");
				return true;
			case Type.InsertPageField:
				this.private_InsertFieldByWordShortcut("PAGE");
				return true;
			case Type.UnlinkField:
				this.private_UnlinkFields();
				return true;
			case Type.NextField:
				this.private_GotoField(true);
				return true;
			case Type.PreviousField:
				this.private_GotoField(false);
				return true;
			case Type.ToggleFieldCode:
			{
				let fields = this.GetCurrentComplexFields();
				if (!fields.length && this.IsSelectionUse())
					fields = this.GetAllFields(true);
				if (fields.length)
				{
					fields[fields.length - 1].ToggleFieldCodes();
					this.UpdateSelection();
				}
				return true;
			}
			case Type.ToggleAllFieldCodes:
				this.private_ToggleAllFieldCodes();
				return true;
			case Type.NextFootnote:
				this.GotoFootnote(true);
				return true;
			case Type.PreviousFootnote:
				this.GotoFootnote(false);
				return true;
			case Type.OptionalHyphen:
				this.private_AddSymbolByShortcut(0x00AD);
				return true;
			case Type.ExtendSelection:
				this.private_ExtendSelectionStep();
				return true;
			case Type.ShrinkSelection:
				this.private_ShrinkSelectionStep();
				return true;
			case Type.MoveToScreenTop:
				this.private_MoveToScreenEdge(false, false);
				return true;
			case Type.MoveToScreenBottom:
				this.private_MoveToScreenEdge(true, false);
				return true;
			case Type.SelectToWindowBottom:
				this.private_MoveToScreenEdge(true, true);
				return true;
			case Type.SelectFloatingObject:
				this.private_SelectFloatingObject();
				return true;
			case Type.RepeatFind:
				if (this.Api.lastFindProps)
					this.Api.asc_findText(this.Api.lastFindProps, true);
				return true;
			case Type.ParagraphRtl:
			case Type.ParagraphLtr:
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetParagraphBidi);
					this.SetParagraphBidi(Type.ParagraphRtl === type);
					this.FinalizeAction();
				}
				return true;
			case Type.InsertTableRowAbove:
				if (getCurrentTableCell(this) && !this.IsSelectionLocked(AscCommon.changestype_Table_Properties))
				{
					this.StartAction(AscDFH.historydescription_Document_TableAddRowAbove);
					this.AddTableRow(true);
					this.FinalizeAction();
				}
				return true;
			case Type.InsertInvertedQuestionMark:
				this.private_AddSymbolByShortcut(0x00BF);
				return true;
			case Type.InsertInvertedExclamationMark:
				this.private_AddSymbolByShortcut(0x00A1);
				return true;
			case Type.NextSpellingError:
				this.private_SelectNextSpellingError();
				return true;
			case Type.SpellingAndGrammar:
				if (this.private_SelectNextSpellingError())
					this.private_OpenContextMenuAtTarget();
				return true;
		}
		return false;
	};

	/**
	 * Alt+X: the hexadecimal code before the caret becomes its character, or the character
	 * before the caret becomes its code (Word). With a selection the editor's own code runs.
	 * @returns {boolean} true when handled
	 */
	CDocument.prototype.private_ToggleUnicode = function()
	{
		if (this.IsSelectionUse())
			return false;

		let paragraph = this.GetCurrentParagraph();
		if (!paragraph || !paragraph.IsParagraph() || docpostype_DrawingObjects === this.GetDocPosType())
			return false;

		let text = paragraph.GetText({ParaEndToSpace : false}).replace(/[\r\n]+$/, "");
		let offset = this.private_GetOffsetInParagraph(paragraph);
		let before = text.substring(0, offset);
		if (!before)
			return false;

		let match = /(?:[Uu]\+)?([0-9A-Fa-f]{2,6})$/.exec(before);
		let replaceLength, insert;
		if (match && parseInt(match[1], 16) > 0x1F && parseInt(match[1], 16) < 0x110000 && !(parseInt(match[1], 16) >= 0xD800 && parseInt(match[1], 16) <= 0xDFFF))
		{
			replaceLength = match[0].length;
			insert = [parseInt(match[1], 16)];
		}
		else
		{
			let code = before.codePointAt(before.length - 1 - (before.length > 1 && /[\uDC00-\uDFFF]/.test(before.charAt(before.length - 1)) ? 1 : 0));
			let hex = code.toString(16).toUpperCase();
			while (hex.length < 4)
				hex = "0" + hex;
			replaceLength = code > 0xFFFF ? 2 : 1;
			insert = hex.split("").map(function(c) { return c.charCodeAt(0); });
		}

		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			return true;

		this.StartAction(AscDFH.historydescription_Document_AddLetter);
		for (let i = 0; i < replaceLength; ++i)
			this.MoveCursorLeft(true, false);
		this.Remove(-1, true, false, false, false);
		for (let i = 0; i < insert.length; ++i)
			this.AddToParagraph(new AscWord.CRunText(insert[i]));
		this.Recalculate();
		this.FinalizeAction();
		return true;
	};
})(window);
