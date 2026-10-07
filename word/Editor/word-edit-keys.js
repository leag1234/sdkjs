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
 * Word editing shortcuts that need a small state machine in the document:
 *   F2 / Shift+F2        move / copy the selection to a place chosen with the cursor, Enter performs, Esc cancels
 *   Ctrl+F3              cut the selection to the Spike (accumulates)
 *   Ctrl+Shift+F3        paste the Spike and clear it
 *   Ctrl+Alt+V           ask the interface for the Paste Special dialog
 *   Shift+F5, Ctrl+Alt+Z go back to the places of the last edits (history of 4 places)
 *   Ctrl+Shift+D / W / H  double and words-only underline, hidden text
 *   Alt+Shift+R          header or footer of the previous section
 * The interface is told through the events asc_onMoveCopyMode(mode: 0 off, 1 move, 2 copy),
 * asc_onSpikeChanged(count, added) and asc_onPasteSpecialDialog.
 */
(function(window)
{
	const Type = Asc.c_oAscDocumentShortcutType;
	const PreventAll = 0xFFFF;
	const MAX_LOCATIONS = 4;
	const STORAGE_PREFIX = "eo-editlocs:";

	function isBodyContent(doc)
	{
		return docpostype_Content === doc.GetDocPosType();
	}

	// ---------------------------------------------------------------- F2 / Shift+F2
	CDocument.prototype.startMoveCopyTo = function(isCopy)
	{
		if (!isBodyContent(this) || !this.IsSelectionUse() || this.IsSelectionEmpty(true))
			return false;

		this.WordMoveCopy = {copy : isCopy, state : this.GetSelectionState(), history : this.History.Index};
		this.sendEvent("asc_onMoveCopyMode", isCopy ? 2 : 1);
		return true;
	};

	CDocument.prototype.cancelMoveCopyTo = function()
	{
		if (!this.WordMoveCopy)
			return;
		this.WordMoveCopy = null;
		this.sendEvent("asc_onMoveCopyMode", 0);
	};

	CDocument.prototype.finishMoveCopyTo = function()
	{
		let mode = this.WordMoveCopy;
		this.cancelMoveCopyTo();
		if (!mode || mode.history !== this.History.Index || !isBodyContent(this))
			return;

		let paragraph = this.GetCurrentParagraph();
		if (!paragraph || !paragraph.IsParagraph())
			return;

		// the destination is where the cursor is; the source selection is selected again for the drag-and-drop code
		let nearPos = paragraph.GetCurrentAnchorPosition();
		this.SetSelectionState(mode.state);
		this.OnEndTextDrag(nearPos, mode.copy);
		this.UpdateInterface();
	};

	// ---------------------------------------------------------------- Spike
	CDocument.prototype.spikeCut = function()
	{
		if (!isBodyContent(this) || this.IsSelectionEmpty(true))
			return false;
		if (this.IsSelectionLocked(AscCommon.changestype_Remove, null, true, this.IsFormFieldEditing()))
			return true;

		this.StartAction(AscDFH.historydescription_Cut);
		let content = this.GetSelectedContent(true);
		this.Remove(-1, true, true, false, false, true);
		this.Recalculate();
		this.UpdateSelection();
		this.FinalizeAction();

		this.WordSpike = this.WordSpike || [];
		this.WordSpike.push(content);
		this.sendEvent("asc_onSpikeChanged", this.WordSpike.length, 1);
		return true;
	};

	CDocument.prototype.spikePaste = function()
	{
		let items = this.WordSpike;
		if (!items || !items.length || !isBodyContent(this))
			return false;
		if (this.IsSelectionUse())
			this.RemoveSelection();
		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			return true;

		this.WordSpike = [];
		this.StartAction(AscDFH.historydescription_Document_PasteHotKey);
		for (let i = 0; i < items.length; ++i)
		{
			let paragraph = this.GetCurrentParagraph();
			if (!paragraph || !paragraph.IsParagraph())
				break;
			items[i].Insert(paragraph.GetCurrentAnchorPosition(), false);
			this.RemoveSelection();
			// Word ends each Spike item with a paragraph mark
			this.AddNewParagraph();
		}
		this.Recalculate();
		this.UpdateSelection();
		this.FinalizeAction();
		this.sendEvent("asc_onSpikeChanged", 0, 0);
		return true;
	};

	// ---------------------------------------------------------------- edit locations
	function getStorage()
	{
		try { return window.localStorage; } catch (e) { return null; }
	}

	CDocument.prototype.private_GetEditLocationKey = function()
	{
		let id = this.Api && (this.Api.documentId || (this.Api.DocInfo && this.Api.DocInfo.get_Id && this.Api.DocInfo.get_Id()));
		return id ? STORAGE_PREFIX + id : null;
	};

	CDocument.prototype.recordEditLocation = function()
	{
		if (!isBodyContent(this) || this.IsActionStarted())
			return;

		let paragraph = this.GetCurrentParagraph();
		if (!paragraph || !paragraph.IsParagraph())
			return;

		let list = this.WordEditLocations = this.WordEditLocations || [];
		let id = paragraph.GetId();
		// consecutive edits in one paragraph are one place; the offset is computed once typing pauses
		if (!list.length || list[0].id !== id)
			list.unshift({id : id, offset : undefined});
		list[0].offset = undefined;
		if (list.length > MAX_LOCATIONS)
			list.length = MAX_LOCATIONS;
		this.WordEditCycle = undefined;

		let me = this;
		clearTimeout(this.WordEditLocationTimer);
		this.WordEditLocationTimer = setTimeout(function() { me.flushEditLocation(); }, 300);
	};

	CDocument.prototype.flushEditLocation = function()
	{
		clearTimeout(this.WordEditLocationTimer);
		let list = this.WordEditLocations;
		if (!list || !list.length || undefined !== list[0].offset || !isBodyContent(this))
			return;

		let paragraph = this.GetCurrentParagraph();
		if (!paragraph || !paragraph.IsParagraph() || paragraph.GetId() !== list[0].id)
		{
			list[0].offset = 0;
			return;
		}
		list[0].offset = this.private_GetOffsetInParagraph(paragraph);

		// the places of the last edits are kept in this browser for the next opening of the document
		let storage = getStorage(), key = this.private_GetEditLocationKey();
		if (storage && key)
		{
			let saved = [];
			for (let i = 0; i < list.length; ++i)
			{
				let index = this.Content.indexOf(AscCommon.g_oTableId.Get_ById(list[i].id));
				if (index >= 0)
					saved.push([index, list[i].offset || 0]);
			}
			try { storage.setItem(key, JSON.stringify(saved)); } catch (e) {}
		}
	};

	CDocument.prototype.private_LoadEditLocations = function()
	{
		let storage = getStorage(), key = this.private_GetEditLocationKey();
		let list = [];
		try
		{
			let saved = storage && key ? JSON.parse(storage.getItem(key) || "[]") : [];
			for (let i = 0; i < saved.length && list.length < MAX_LOCATIONS; ++i)
			{
				let element = this.Content[saved[i][0]];
				if (element && element.IsParagraph())
					list.push({id : element.GetId(), offset : saved[i][1]});
			}
		}
		catch (e) {}
		return list;
	};

	CDocument.prototype.private_GoToEditLocation = function(location)
	{
		let paragraph = AscCommon.g_oTableId.Get_ById(location.id);
		if (!paragraph || !paragraph.IsParagraph() || !paragraph.IsUseInDocument())
			return false;

		this.RemoveSelection();
		paragraph.MoveCursorToStartPos(false);
		for (let i = 0; i < (location.offset || 0); ++i)
			paragraph.MoveCursorRight(false, false);
		paragraph.Document_SetThisElementCurrent(true);
		this.UpdateSelection();
		this.UpdateInterface();
		return true;
	};

	/**
	 * Shift+F5, Ctrl+Alt+Z: the place of the last edit, then of the edit before it...
	 * Right after opening the document: the place of the last edit of the previous session, if this browser
	 * remembers it, else the previous tracked change.
	 */
	CDocument.prototype.goToPreviousEditLocation = function()
	{
		if (!isBodyContent(this))
			return false;

		this.flushEditLocation();
		let list = this.WordEditLocations;
		if (!list || !list.length)
		{
			list = this.WordEditLocations = this.private_LoadEditLocations();
			this.WordEditCycle = undefined;
			if (!list.length)
			{
				this.GetPrevRevisionChange();
				return true;
			}
		}

		let paragraph = this.GetCurrentParagraph();
		let index = this.WordEditCycle;
		// the cycle goes on only while the cursor is where the previous jump put it
		if (undefined !== index && (!paragraph || paragraph.GetId() !== this.WordEditCycleId))
			index = undefined;
		if (undefined === index)
			index = (list.length > 1 && paragraph && !this.IsSelectionUse() && list[0].id === paragraph.GetId()) ? 1 : 0;

		for (let tries = 0; tries < list.length; ++tries)
		{
			let location = list[index % list.length];
			index = (index + 1) % list.length;
			if (this.private_GoToEditLocation(location))
			{
				this.WordEditCycleId = location.id;
				break;
			}
		}
		this.WordEditCycle = index;
		return true;
	};

	// FinalizeAction closes every editing action: the places of the edits are recorded there
	let finalizeAction = CDocument.prototype.FinalizeAction;
	CDocument.prototype.FinalizeAction = function(checkEmptyAction, additional)
	{
		let isOuter = this.IsActionStarted() && 0 === this.Action.Depth && this.Action.PointsCount > 0 && !this.Action.CancelAction;
		let result  = finalizeAction.call(this, checkEmptyAction, additional);
		if (isOuter && false !== result && !this.IsActionStarted())
			this.recordEditLocation();
		return result;
	};

	// ---------------------------------------------------------------- underline styles
	/**
	 * Ctrl+Shift+D (double), Ctrl+Shift+W (words only): apply the style, or remove the underline when the text has it
	 */
	CDocument.prototype.private_ToggleUnderlineType = function(type)
	{
		let textPr = this.GetCalculatedTextPr();
		if (!textPr || this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			return;

		let value = textPr.Underline === type ? false : type;
		this.StartAction(AscDFH.historydescription_Document_SetTextUnderlineHotKey, null, null, !!value);
		this.AddToParagraph(new ParaTextPr({Underline : value}));
		this.UpdateInterface();
		this.FinalizeAction();
	};

	// ---------------------------------------------------------------- hidden text
	/**
	 * Ctrl+Shift+H: hide the selected text, or show it again when it is hidden
	 */
	CDocument.prototype.private_ToggleHiddenText = function()
	{
		let textPr = this.GetCalculatedTextPr();
		if (!textPr || this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			return;

		let value = !textPr.Vanish;
		this.StartAction(AscDFH.historydescription_Document_SetTextVanish ? AscDFH.historydescription_Document_SetTextVanish : AscDFH.historydescription_Document_SetTextUnderlineHotKey, null, null, value);
		this.AddToParagraph(new ParaTextPr({Vanish : value}));
		this.UpdateInterface();
		this.FinalizeAction();
	};

	/**
	 * The formatting marks were switched: the hidden text appears or disappears, the paragraphs that have some
	 * are measured and laid out again
	 */
	CDocument.prototype.refreshHiddenText = function()
	{
		let changed = false;
		let paragraphs = this.GetAllParagraphs({All : true});
		for (let i = 0; i < paragraphs.length; ++i)
		{
			let hasHidden = false;
			paragraphs[i].CheckRunContent(function(run)
			{
				if (run.Get_CompiledPr(false).Vanish)
					hasHidden = true;
				return hasHidden;
			});
			if (hasHidden)
			{
				paragraphs[i].Recalc_RunsCompiledPr();
				changed = true;
			}
		}
		if (changed)
			this.RecalculateFromStart(true);
	};

	// ---------------------------------------------------------------- shortcuts
	let executeWordShortcut = CDocument.prototype.executeWordShortcut;
	CDocument.prototype.executeWordShortcut = function(type)
	{
		switch (type)
		{
			case Type.MoveToLocation:
				this.startMoveCopyTo(false);
				return true;
			case Type.CopyToLocation:
				this.startMoveCopyTo(true);
				return true;
			case Type.SpikeCut:
				this.spikeCut();
				return true;
			case Type.SpikePaste:
				this.spikePaste();
				return true;
			case Type.PasteSpecial:
				// Windows: Ctrl+Alt+V first applies copied formatting (Ctrl+Alt+C), as it did before, then it is Paste Special
				if (!AscCommon.AscBrowser.isMacOs && this.Api.getFormatPainterData && this.Api.getFormatPainterData())
				{
					if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
					{
						this.StartAction(AscDFH.historydescription_Document_FormatPasteHotKey);
						this.Document_Format_Paste();
						this.FinalizeAction();
					}
					this.Api.clearFormatPainterData();
					this.Api.sendPaintFormatEvent(false);
					return true;
				}
				this.sendEvent("asc_onPasteSpecialDialog");
				return true;
			case Type.UnderlineDouble:
				this.private_ToggleUnderlineType(Asc.UnderlineType.Double);
				return true;
			case Type.UnderlineWords:
				this.private_ToggleUnderlineType(Asc.UnderlineType.Words);
				return true;
			case Type.CopyPreviousHeaderFooter:
				// the header or footer being edited becomes the one of the previous section (Link to Previous)
				if (docpostype_HdrFtr === this.GetDocPosType())
					this.Api.HeadersAndFooters_LinkToPrevious(true);
				return true;
			case Type.HiddenText:
				this.private_ToggleHiddenText();
				return true;
			case Type.PreviousEditLocation:
			case Type.CycleEditLocations:
				this.goToPreviousEditLocation();
				return true;
		}
		return executeWordShortcut ? executeWordShortcut.call(this, type) : false;
	};

	// keys while a move / copy is pending: Enter performs, Esc cancels, navigation keys choose the place
	let handleWordKeyDown = CDocument.prototype.handleWordKeyDown;
	CDocument.prototype.handleWordKeyDown = function(e)
	{
		if (this.WordMoveCopy)
		{
			let code = e.KeyCode;
			let isModifier = 16 === code || 17 === code || 18 === code || 91 === code || 93 === code || 224 === code;
			let isNavigation = (code >= 33 && code <= 40) || 9 === code;
			if (13 === code && !e.CtrlKey && !e.AltKey && !e.ShiftKey)
			{
				this.finishMoveCopyTo();
				return PreventAll;
			}
			if (27 === code)
			{
				this.cancelMoveCopyTo();
				return PreventAll;
			}
			if (113 === code)
				return PreventAll;
			if (!isModifier && !isNavigation)
				this.cancelMoveCopyTo();
		}
		return handleWordKeyDown ? handleWordKeyDown.call(this, e) : undefined;
	};
})(window);
