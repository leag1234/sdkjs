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
 * Microsoft Excel keyboard behaviour that the spreadsheet editor lacked:
 *  - Repeat (F4 outside the cell editor, Ctrl+Y / Cmd+Y when there is nothing to redo): the last
 *    formatting, insert, delete, row height or column width command is applied again to the
 *    current selection. It is recorded at WorksheetView.setSelectionInfo / changeWorksheet with
 *    the value it applied (Ctrl+B is recorded as "bold on", not as a toggle);
 *  - End mode (End, then an arrow key or Home), Extend Selection (F8) and Add to Selection
 *    (Shift+F8), announced with the event asc_onKeyboardSelectionMode ("end", "extend", "add", "");
 *  - Ctrl+Period (next corner of the selection), Ctrl+Shift+* (current region), Shift+Backspace
 *    (active cell only), Alt+Page Down / Page Up (one screen right / left), Ctrl+' and Ctrl+Shift+"
 *    (formula or value of the cell above), Ctrl+Alt+5 (first floating object), Ctrl+E (Flash Fill);
 *  - cell editor: F2 switches between Enter (point) and Edit mode, Ctrl+Delete deletes to the end
 *    of the line, Ctrl+A right of a function name opens its Function Arguments dialog, Cmd+T cycles
 *    the reference type on macOS like F4;
 *  - macOS: Option+Right/Left (next/previous sheet), Delete (Backspace) on a range clears it,
 *    Ctrl+U edits the active cell.
 */
(function (window) {
	const AscCommonExcel = window["AscCommonExcel"];
	const EventsController = AscCommonExcel.asc_CEventsController;
	const CellEditor = AscCommonExcel.CellEditor;
	const WorksheetView = AscCommonExcel.WorksheetView;
	const WorkbookView = AscCommonExcel.WorkbookView;

	const isMac = function () {
		return AscCommon.AscBrowser.isMacOs;
	};
	// The physical Control key (on macOS CtrlKey is also set by Cmd)
	const physCtrl = function (e) {
		return isMac() ? !!e.MacCtrlKey : !!e.CtrlKey && !e.AltGr;
	};
	const noMods = function (e) {
		return !e.CtrlKey && !e.AltKey && !e.ShiftKey && !e.MacCmdKey;
	};
	const cloneEvent = function (e, patch) {
		const copy = Object.assign(Object.create(Object.getPrototypeOf(e)), e);
		return Object.assign(copy, patch);
	};

	// ---- Repeat --------------------------------------------------------------------------------

	const REPEATABLE_SELECTION_PROPS = {
		"fn": 1, "fs": 1, "b": 1, "i": 1, "u": 1, "s": 1, "fa": 1, "a": 1, "readingOrder": 1, "va": 1, "c": 1,
		"f": 1, "bc": 1, "wrap": 1, "shrink": 1, "format": 1, "angle": 1, "indent": 1, "locked": 1,
		"hiddenFormulas": 1, "border": 1, "merge": 1, "empty": 1, "changeDigNum": 1, "changeFontSize": 1,
		"style": 1, "changeTextCase": 1
	};
	const REPEATABLE_WORKSHEET_PROPS = {
		"insCell": 1, "delCell": 1, "colWidth": 1, "rowHeight": 1, "hideCols": 1, "hideRows": 1, "showCols": 1, "showRows": 1
	};

	const repeatState = {last: null, replaying: false};

	function recordRepeat(kind, prop, args) {
		if (repeatState.replaying) {
			return;
		}
		const api = window["Asc"]["editor"];
		if (!api || !api.canEdit()) {
			return;
		}
		repeatState.last = {kind: kind, prop: prop, args: Array.prototype.slice.call(args, 0, 2)};
	}

	function executeRepeat(wb) {
		const last = repeatState.last;
		if (!last || wb.getCellEditMode() || !wb.Api.canEdit()) {
			return false;
		}
		const ws = wb.getWorksheet();
		repeatState.replaying = true;
		try {
			if ("selection" === last.kind) {
				ws.setSelectionInfo(last.prop, last.args[1]);
			} else {
				ws.changeWorksheet(last.prop, last.args[1]);
			}
		} finally {
			repeatState.replaying = false;
		}
		return true;
	}

	const originalSetSelectionInfo = WorksheetView.prototype.setSelectionInfo;
	WorksheetView.prototype.setSelectionInfo = function (prop, val) {
		if (REPEATABLE_SELECTION_PROPS[prop]) {
			recordRepeat("selection", prop, arguments);
		}
		return originalSetSelectionInfo.apply(this, arguments);
	};
	const originalChangeWorksheet = WorksheetView.prototype.changeWorksheet;
	WorksheetView.prototype.changeWorksheet = function (prop, val) {
		if (REPEATABLE_WORKSHEET_PROPS[prop]) {
			recordRepeat("worksheet", prop, arguments);
		}
		return originalChangeWorksheet.apply(this, arguments);
	};
	// Ctrl+Y repeats when there is nothing to redo, as in Excel
	const originalRedo = WorkbookView.prototype.redo;
	WorkbookView.prototype.redo = function () {
		if (!this.getCellEditMode() && !this.Api.isEditVisibleAreaOleEditor && !AscCommon.History.Can_Redo() && repeatState.last) {
			return executeRepeat(this);
		}
		return originalRedo.apply(this, arguments);
	};
	WorkbookView.prototype.asc_Repeat = function () {
		return executeRepeat(this);
	};
	WorkbookView.prototype.canRepeat = function () {
		return !!repeatState.last;
	};

	// ---- Keyboard selection modes ----------------------------------------------------------------

	function setMode(controller, mode) {
		if (controller.elixirKeyMode === mode) {
			return;
		}
		controller.elixirKeyMode = mode;
		controller.elixirAddStarted = false;
		const api = controller.view && controller.view.Api;
		if (api) {
			api.sendEvent("asc_onKeyboardSelectionMode", mode || "");
		}
	}

	function arrowDelta(code) {
		switch (code) {
			case 37: return [-1, 0];
			case 38: return [0, -1];
			case 39: return [1, 0];
			case 40: return [0, 1];
		}
		return null;
	}

	function redrawSelection(ws) {
		ws.updateSelectionWithSparklines();
		ws._updateSelectionNameAndInfo();
		ws._scrollToRange && ws._scrollToRange();
	}

	// Ctrl+Period: the active cell goes to the next corner of the selection, clockwise
	function nextCorner(ws) {
		const sr = ws.model.selectionRange;
		const r = sr.getLast();
		if (r.isOneCell()) {
			return;
		}
		const a = sr.activeCell;
		let row = r.r1, col = r.c2;
		if (a.row === r.r1 && a.col === r.c2) {
			row = r.r2; col = r.c2;
		} else if (a.row === r.r2 && a.col === r.c2) {
			row = r.r2; col = r.c1;
		} else if (a.row === r.r2 && a.col === r.c1) {
			row = r.r1; col = r.c1;
		}
		ws.cleanSelection();
		a.row = row;
		a.col = col;
		redrawSelection(ws);
	}

	// Ctrl+Shift+*: select the current region around the active cell
	function selectCurrentRegion(ws) {
		const sr = ws.model.selectionRange;
		const a = sr.activeCell.clone();
		const region = ws.model.autoFilters.expandRange(new Asc.Range(a.col, a.row, a.col, a.row), true, true, true);
		if (!region) {
			return;
		}
		ws.cleanSelection();
		sr.clean();
		sr.getLast().assign2(region);
		sr.activeCell.row = a.row;
		sr.activeCell.col = a.col;
		sr.update && sr.update();
		redrawSelection(ws);
	}

	// Shift+Backspace: keep only the active cell
	function collapseToActive(ws) {
		const sr = ws.model.selectionRange;
		const a = sr.activeCell.clone();
		ws.cleanSelection();
		sr.clean();
		sr.getLast().assign(a.col, a.row, a.col, a.row);
		sr.activeCell.row = a.row;
		sr.activeCell.col = a.col;
		redrawSelection(ws);
	}

	// Ctrl+' / Ctrl+Shift+": the formula (or the value) of the cell above
	function textOfCellAbove(ws, value) {
		const sr = ws.model.selectionRange || ws.model.copySelection;
		const a = sr && sr.activeCell;
		if (!a || a.row < 1) {
			return "";
		}
		const range = ws.model.getRange3(a.row - 1, a.col, a.row - 1, a.col);
		return value ? range.getValueWithFormat() : range.getValueForEdit();
	}

	function selectFirstObject(ws) {
		const render = ws.objectRender;
		const ctrl = render && render.controller;
		const drawings = ctrl && ctrl.getDrawingObjects();
		if (!drawings || !drawings.length) {
			return false;
		}
		ctrl.resetSelection();
		drawings[0].select(ctrl, 0);
		ctrl.updateSelectionState && ctrl.updateSelectionState();
		render.showDrawingObjects && render.showDrawingObjects();
		return true;
	}

	/**
	 * Keys of the grid. Returns a keydown result when the key was handled here, undefined otherwise.
	 * @param controller {asc_CEventsController}
	 * @param e {AscCommon.CKeyboardEvent}
	 */
	function onGridKey(controller, e) {
		if (!controller.enableKeyEvents || controller.getCellEditMode() || controller.getSelectionDialogMode() ||
			controller.isSelectMode || controller.isFillHandleMode || controller.isMoveRangeMode || controller.isMoveResizeRange) {
			return undefined;
		}
		const wb = controller.view;
		const ws = wb.getWorksheet();
		const code = e.GetKeyCode();
		const canEdit = controller.canEdit();
		const mac = isMac();
		const mode = controller.elixirKeyMode;

		if (16 === code || 17 === code || 18 === code || 91 === code || 93 === code || 224 === code) {
			return undefined;
		}

		// F4: Repeat
		if (115 === code && noMods(e)) {
			if (canEdit) {
				executeRepeat(wb);
			}
			return keydownresult_PreventAll;
		}
		// F8: Extend Selection; Shift+F8: Add to Selection
		if (119 === code && !e.CtrlKey && !e.AltKey && !e.MacCmdKey) {
			setMode(controller, e.ShiftKey ? ("add" === mode ? null : "add") : ("extend" === mode ? null : "extend"));
			return keydownresult_PreventAll;
		}
		// End mode (Windows keyboards; End alone does not move the active cell)
		if (35 === code && noMods(e) && !mac) {
			setMode(controller, "end" === mode ? null : "end");
			return keydownresult_PreventAll;
		}
		if (27 === code && mode) {
			setMode(controller, null);
			return undefined;
		}
		if ("end" === mode) {
			setMode(controller, null);
			if (arrowDelta(code) && !e.CtrlKey && !e.AltKey) {
				return controller.elixirOriginalKeyDown(cloneEvent(e, {CtrlKey: true}));
			}
			if (36 === code && !e.CtrlKey && !e.AltKey) {
				return controller.elixirOriginalKeyDown(cloneEvent(e, {CtrlKey: true, KeyCode: 35}));
			}
			return undefined;
		}
		if ("extend" === mode && !e.CtrlKey && !e.AltKey && !e.ShiftKey && (arrowDelta(code) || 33 === code || 34 === code || 36 === code)) {
			return controller.elixirOriginalKeyDown(cloneEvent(e, {ShiftKey: true}));
		}
		if ("add" === mode && arrowDelta(code) && !e.CtrlKey && !e.AltKey && !e.ShiftKey) {
			const d = arrowDelta(code);
			const sr = ws.model.selectionRange;
			if (controller.elixirAddStarted && sr.ranges.length > 1) {
				sr.ranges.pop();
				sr.activeCellId = sr.ranges.length - 1;
			}
			controller.elixirAddStarted = true;
			controller.handlers.trigger("changeSelection", true, d[0], d[1], false, true, function (delta) {
				controller.scroll(ws.convertOffsetToSmooth(delta));
			});
			return keydownresult_PreventAll;
		}
		if (mode && "add" !== mode && !(e.ShiftKey && 16 === code)) {
			// any other key ends Extend mode
			if (!arrowDelta(code)) {
				setMode(controller, null);
			}
		}

		// Ctrl+Period: next corner of the selection
		if (190 === code && physCtrl(e) && !e.ShiftKey && !e.AltKey && !e.MacCmdKey) {
			nextCorner(ws);
			return keydownresult_PreventAll;
		}
		// Ctrl+Shift+* (Ctrl+Shift+8, Ctrl+numpad *): current region
		if (!mac && physCtrl(e) && !e.AltKey && ((56 === code && e.ShiftKey) || 106 === code)) {
			selectCurrentRegion(ws);
			return keydownresult_PreventAll;
		}
		// Shift+Backspace: only the active cell
		if (8 === code && e.ShiftKey && !e.CtrlKey && !e.AltKey) {
			collapseToActive(ws);
			return keydownresult_PreventAll;
		}
		// Alt+Page Down / Page Up: one screen right / left
		if ((33 === code || 34 === code) && e.AltKey && !e.CtrlKey && !e.MacCmdKey) {
			controller.handlers.trigger("changeSelection", !e.ShiftKey, 34 === code ? 0.5 : -0.5, 0, false, false, function (delta) {
				controller.scroll(ws.convertOffsetToSmooth(delta));
			});
			return keydownresult_PreventAll;
		}
		// Ctrl+' (formula above), Ctrl+Shift+" (value above)
		if (222 === code && physCtrl(e) && !e.AltKey && !e.MacCmdKey && canEdit) {
			const options = new AscCommonExcel.CEditorEnterOptions();
			options.newText = textOfCellAbove(ws, e.ShiftKey);
			options.focus = true;
			controller.handlers.trigger("editCell", options);
			controller.setSkipCellEditor(true);
			return keydownresult_PreventAll;
		}
		// Ctrl+E: Flash Fill (cell/model/flash-fill.js)
		if (69 === code && physCtrl(e) && !e.ShiftKey && !e.AltKey && !e.MacCmdKey && canEdit) {
			wb.Api.asc_FlashFill();
			return keydownresult_PreventAll;
		}
		// Ctrl+Alt+5: first floating object (then Tab cycles through them)
		if (53 === code && !mac && (e.AltGr || (e.CtrlKey && e.AltKey)) && !e.ShiftKey) {
			selectFirstObject(ws);
			return keydownresult_PreventAll;
		}
		if (mac) {
			// Option+Right / Left: next / previous sheet
			if ((37 === code || 39 === code) && e.AltKey && !e.CtrlKey && !e.ShiftKey) {
				controller.handlers.trigger("showNextPrevWorksheet", 39 === code ? +1 : -1);
				return keydownresult_PreventAll;
			}
			// Delete (Backspace) on a range clears it, on one cell it edits the cell (existing behaviour)
			if (8 === code && noMods(e) && canEdit && (ws.model.selectionRange.ranges.length > 1 || !ws.model.selectionRange.getLast().isOneCell())) {
				controller.handlers.trigger("empty");
				return keydownresult_PreventAll;
			}
			// Ctrl+U: edit the active cell, insertion point at the end
			if (85 === code && e.MacCtrlKey && !e.MacCmdKey && !e.AltKey && !e.ShiftKey && canEdit) {
				const options = new AscCommonExcel.CEditorEnterOptions();
				options.focus = true;
				controller.handlers.trigger("editCell", options);
				controller.setSkipCellEditor(true);
				return keydownresult_PreventAll;
			}
		}
		return undefined;
	}

	const originalControllerKeyDown = EventsController.prototype._onWindowKeyDown;
	EventsController.prototype.elixirOriginalKeyDown = function (e) {
		return originalControllerKeyDown.call(this, e);
	};
	EventsController.prototype._onWindowKeyDown = function (e) {
		const res = onGridKey(this, e);
		if (undefined !== res) {
			if (res & keydownresult_PreventKeyPress) {
				this._setSkipKeyPress(true);
			}
			return res;
		}
		return originalControllerKeyDown.apply(this, arguments);
	};
	// A mouse selection ends End and Add modes, as in Excel
	const originalMouseDown = EventsController.prototype._onMouseDown;
	if (originalMouseDown) {
		EventsController.prototype._onMouseDown = function () {
			if ("end" === this.elixirKeyMode) {
				setMode(this, null);
			}
			return originalMouseDown.apply(this, arguments);
		};
	}

	// The cell editor keys below only apply when the editor was already open before the key (F2 in the
	// grid opens the editor, then the same key event reaches the editor)
	const originalWbKeyDown = WorkbookView.prototype.onKeyDown;
	WorkbookView.prototype.onKeyDown = function (e) {
		this.cellEditor.elixirOpenBeforeKey = !!this.getCellEditMode();
		return originalWbKeyDown.apply(this, arguments);
	};

	// ---- Cell editor ---------------------------------------------------------------------------

	const kEndOfLine = -3;
	const kNextChar = -5;

	const originalEditorKeyDown = CellEditor.prototype._onWindowKeyDown;
	CellEditor.prototype._onWindowKeyDown = function (e) {
		if (this.isOpened && this.elixirOpenBeforeKey && !this.handlers.trigger("getWizard") && this.enableKeyEvents) {
			const code = e.GetKeyCode();
			// F2: Enter (point) mode <-> Edit mode
			if (113 === code && noMods(e)) {
				this.setFocus(!this.hasFocus);
				return keydownresult_PreventAll;
			}
			// Ctrl+Delete: delete to the end of the line
			if (46 === code && physCtrl(e) && !e.ShiftKey && !e.AltKey && !e.MacCmdKey) {
				this._selectChars(kEndOfLine);
				this._removeChars(kNextChar);
				this._setSkipKeyPress(true);
				return keydownresult_PreventAll;
			}
			// Ctrl+' / Ctrl+Shift+" in the editor: insert the formula or value of the cell above
			if (222 === code && physCtrl(e) && !e.AltKey && !e.MacCmdKey) {
				const wb = window["Asc"]["editor"].wb;
				const text = textOfCellAbove(wb.getWorksheet(), e.ShiftKey);
				if (text) {
					this._addChars(text);
				}
				this._setSkipKeyPress(true);
				return keydownresult_PreventAll;
			}
			// Ctrl+A right of a function name: the Function Arguments dialog of that function
			if (65 === code && physCtrl(e) && !e.ShiftKey && !e.AltKey && !e.MacCmdKey && this.isFormula()) {
				const api = window["Asc"]["editor"];
				const match = /([A-Za-z][A-Za-z0-9._]*)\(?$/.exec(this.getText(0, this.cursorPos));
				const name = match && api.asc_getFormulaNameByLocale(match[1].toUpperCase());
				if (name && AscCommonExcel.cFormulaFunction[name]) {
					this._removeChars(this.cursorPos - match[0].length, match[0].length);
					api.asc_startWizard(name, false);
					this._setSkipKeyPress(true);
					return keydownresult_PreventAll;
				}
			}
			// Cmd+T (macOS): cycle the reference type, like F4
			if (isMac() && 84 === code && e.MacCmdKey && !e.MacCtrlKey && !e.AltKey && !e.ShiftKey) {
				const res = this._findRangeUnderCursor();
				if (res.range) {
					res.range.switchReference();
					this.changeCellRange(res.range);
				}
				this._setSkipKeyPress(true);
				return keydownresult_PreventAll;
			}
		}
		return originalEditorKeyDown.apply(this, arguments);
	};
})(window);
