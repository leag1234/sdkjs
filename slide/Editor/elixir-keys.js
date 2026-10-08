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
 * Microsoft PowerPoint keyboard behaviour that the presentation editor lacked:
 *  - in a text box: Ctrl+Up / Ctrl+Down (Cmd on macOS) to the start of the previous / next
 *    paragraph (Ctrl+Up first goes to the start of the current one), Ctrl+Shift+Up / Down to
 *    select to the start / end of the paragraph, Alt+Shift+Up / Down (Ctrl+Shift+Up / Down on
 *    macOS) to move the selected paragraphs up or down;
 *  - Ctrl+Alt+5: select the first floating object of the slide (then Tab cycles through them);
 *  - Ctrl+Shift+J: regroup the objects of the last ungrouped group;
 *  - Shift+F4: repeat the last Find;
 *  - asc_GetPaneFocus / asc_SetPaneFocus for F6 and Shift+F6 (slide, thumbnails) in the interface.
 */
(function (window) {
	const CPresentation = window["AscCommonSlide"].CPresentation;
	const api = window["Asc"]["asc_docs_api"];

	const isMac = function () {
		return AscCommon.AscBrowser.isMacOs;
	};
	// Ctrl on Windows and Linux, Cmd on macOS
	const primary = function (e) {
		return isMac() ? !!e.MacCmdKey && !e.MacCtrlKey : !!e.CtrlKey && !e.AltGr;
	};

	function getTextContent(presentation) {
		const controller = presentation.GetCurrentController();
		const content = controller && controller.getTargetDocContent();
		return content && content.Content && content.Content.length ? {controller: controller, content: content} : null;
	}

	function currentParagraphIndex(content) {
		if (content.IsSelectionUse()) {
			return content.Selection.EndPos;
		}
		return content.CurPos.ContentPos;
	}

	function paragraphEdge(content, toEnd) {
		const index = currentParagraphIndex(content);
		const para = content.Content[index];
		if (!para || !para.IsParagraph || !para.IsParagraph()) {
			return true;
		}
		const pos = para.Get_ParaContentPos(content.IsSelectionUse(), false);
		return toEnd ? para.IsCursorAtEnd(pos) : para.IsCursorAtBegin(pos);
	}

	const MAX_STEPS = 20000;

	// Ctrl+Up / Ctrl+Down: start of the current, previous or next paragraph
	function moveByParagraph(presentation, text, down) {
		const controller = text.controller, content = text.content;
		if (content.IsSelectionUse()) {
			down ? controller.cursorMoveRight(false, false) : controller.cursorMoveLeft(false, false);
		}
		const start = currentParagraphIndex(content);
		if (down) {
			if (start >= content.Content.length - 1) {
				controller.cursorMoveToEndPos && controller.cursorMoveToEndPos(false);
			} else {
				for (let i = 0; i < MAX_STEPS && currentParagraphIndex(content) === start; ++i) {
					controller.cursorMoveRight(false, false);
				}
			}
		} else {
			if (paragraphEdge(content, false) && start > 0) {
				controller.cursorMoveLeft(false, false);
			}
			for (let i = 0; i < MAX_STEPS && !paragraphEdge(content, false); ++i) {
				controller.cursorMoveLeft(false, false);
			}
		}
	}

	// Ctrl+Shift+Up / Down: extend the selection to the start / end of the paragraph
	function selectToParagraphEdge(presentation, text, down) {
		const controller = text.controller, content = text.content;
		if (paragraphEdge(content, down)) {
			down ? controller.cursorMoveRight(true, false) : controller.cursorMoveLeft(true, false);
		}
		const index = currentParagraphIndex(content);
		for (let i = 0; i < MAX_STEPS && !paragraphEdge(content, down); ++i) {
			down ? controller.cursorMoveRight(true, false) : controller.cursorMoveLeft(true, false);
			if (currentParagraphIndex(content) !== index) {
				down ? controller.cursorMoveLeft(true, false) : controller.cursorMoveRight(true, false);
				break;
			}
		}
	}

	// Alt+Shift+Up / Down: move the paragraphs of the selection over the previous / next one
	function moveParagraphs(presentation, text, down) {
		const content = text.content;
		let first, last;
		if (content.IsSelectionUse()) {
			first = Math.min(content.Selection.StartPos, content.Selection.EndPos);
			last = Math.max(content.Selection.StartPos, content.Selection.EndPos);
		} else {
			first = last = content.CurPos.ContentPos;
		}
		const other = down ? last + 1 : first - 1;
		if (other < 0 || other >= content.Content.length) {
			return;
		}
		if (false !== presentation.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props)) {
			return;
		}
		const selection = content.GetSelectionState ? content.GetSelectionState() : null;
		presentation.StartAction(AscDFH.historydescription_Document_MoveParagraphs || AscDFH.historydescription_Presentation_ParagraphAdd);
		const moved = content.Content[other];
		content.Internal_Content_Remove(other, 1);
		content.Internal_Content_Add(down ? first : last, moved);
		if (selection && content.SetSelectionState) {
			const shift = down ? 1 : -1;
			const shiftState = function (state) {
				if (state && typeof state === "object") {
					["StartPos", "EndPos", "ContentPos"].forEach(function (k) {
						if (typeof state[k] === "number") {
							state[k] += shift;
						}
					});
				}
			};
			// the last element of the state describes this content: its positions move with the paragraphs
			const own = selection[selection.length - 1];
			shiftState(own && own.Selection);
			shiftState(own && own.CurPos);
			content.SetSelectionState(selection, selection.length - 1);
		}
		presentation.Recalculate();
		presentation.Document_UpdateSelectionState();
		presentation.Document_UpdateInterfaceState();
		presentation.FinalizeAction();
	}

	function selectFirstObject(presentation) {
		const slide = presentation.GetCurrentSlide();
		const controller = presentation.GetCurrentController();
		if (!slide || !controller || !slide.cSld.spTree.length) {
			return false;
		}
		controller.resetSelection();
		slide.cSld.spTree[0].select(controller, presentation.CurPage);
		presentation.Document_UpdateSelectionState();
		presentation.Document_UpdateInterfaceState();
		return true;
	}

	// ---- Regroup --------------------------------------------------------------------------------
	const originalUnGroup = CPresentation.prototype.unGroupShapes;
	CPresentation.prototype.unGroupShapes = function () {
		const res = originalUnGroup.apply(this, arguments);
		const controller = this.GetCurrentController();
		const slide = this.GetCurrentSlide();
		if (controller && slide && controller.selectedObjects.length > 1) {
			this.elixirLastUngrouped = {slide: slide, ids: controller.selectedObjects.map(function (o) {
				return o.Get_Id();
			})};
		}
		return res;
	};
	CPresentation.prototype.elixirRegroup = function () {
		const last = this.elixirLastUngrouped;
		const slide = this.GetCurrentSlide();
		const controller = this.GetCurrentController();
		if (!last || !slide || last.slide !== slide || !controller || !this.CanEdit()) {
			return false;
		}
		const objects = last.ids.map(function (id) {
			return AscCommon.g_oTableId.Get_ById(id);
		}).filter(function (o) {
			return o && slide.cSld.spTree.indexOf(o) >= 0;
		});
		if (objects.length < 2) {
			return false;
		}
		controller.resetSelection();
		const page = this.CurPage;
		objects.forEach(function (o) {
			o.select(controller, page);
		});
		this.groupShapes();
		this.elixirLastUngrouped = null;
		return true;
	};

	// ---- Repeat the last Find (Shift+F4) ---------------------------------------------------------
	const originalFindText = api.prototype.asc_findText;
	api.prototype.asc_findText = function (props) {
		if (props) {
			this.elixirLastFind = props;
		}
		return originalFindText.apply(this, arguments);
	};

	// ---- Pane focus (F6 / Shift+F6 in the interface) ----------------------------------------------
	/**
	 * @returns {string} "thumbnails" or "slide": where the keyboard goes in the editor
	 */
	api.prototype.asc_GetPaneFocus = function () {
		const thumbnails = this.WordControl && this.WordControl.Thumbnails;
		return thumbnails && thumbnails.FocusObjType === AscCommon.FOCUS_OBJECT_THUMBNAILS ? "thumbnails" : "slide";
	};
	/**
	 * @param pane {string} "thumbnails" or "slide"
	 */
	api.prototype.asc_SetPaneFocus = function (pane) {
		const thumbnails = this.WordControl && this.WordControl.Thumbnails;
		if (!thumbnails) {
			return;
		}
		thumbnails.SetFocusElement("thumbnails" === pane ? AscCommon.FOCUS_OBJECT_THUMBNAILS : AscCommon.FOCUS_OBJECT_MAIN);
		this.asc_enableKeyEvents(true);
	};
	api.prototype["asc_GetPaneFocus"] = api.prototype.asc_GetPaneFocus;
	api.prototype["asc_SetPaneFocus"] = api.prototype.asc_SetPaneFocus;

	// ---- Keys ------------------------------------------------------------------------------------
	const originalOnKeyDown = CPresentation.prototype.OnKeyDown;
	CPresentation.prototype.OnKeyDown = function (e) {
		const code = e.KeyCode;
		const mac = isMac();
		if (38 === code || 40 === code) {
			const down = 40 === code;
			const text = getTextContent(this);
			if (text && !e.AltKey && primary(e)) {
				if (e.ShiftKey) {
					selectToParagraphEdge(this, text, down);
				} else {
					moveByParagraph(this, text, down);
				}
				this.private_UpdateCursorXY(true, true);
				this.Document_UpdateSelectionState();
				this.Document_UpdateInterfaceState();
				return keydownresult_PreventAll;
			}
			const moveKeys = mac ? (e.MacCtrlKey && !e.MacCmdKey && !e.AltKey && e.ShiftKey) : (e.AltKey && e.ShiftKey && !e.CtrlKey);
			if (text && moveKeys && this.CanEdit()) {
				moveParagraphs(this, text, down);
				return keydownresult_PreventAll;
			}
		}
		// Ctrl+Alt+5 (Windows: AltGr is Ctrl+Alt)
		if (53 === code && !mac && (e.AltGr || (e.CtrlKey && e.AltKey)) && !e.ShiftKey) {
			selectFirstObject(this);
			return keydownresult_PreventAll;
		}
		if (74 === code && primary(e) && e.ShiftKey && !e.AltKey) {
			this.elixirRegroup();
			return keydownresult_PreventAll;
		}
		if (115 === code && e.ShiftKey && !e.CtrlKey && !e.AltKey && !e.MacCmdKey) {
			if (this.Api.elixirLastFind) {
				this.Api.asc_findText(this.Api.elixirLastFind, true);
			}
			return keydownresult_PreventAll;
		}
		return originalOnKeyDown.apply(this, arguments);
	};
})(window);
