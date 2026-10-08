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
 * Quick Analysis, as in Microsoft Excel: when a range with data is selected, the interface shows a
 * button at its bottom-right corner (event asc_onQuickAnalysis with the corner position, or null)
 * and Ctrl+Q opens the same gallery. The formatting and chart commands reuse the existing API
 * (asc_setCF, asc_addChartDrawingObject...); this file adds the position and the Totals command.
 */
(function (window) {
	const AscCommonExcel = window["AscCommonExcel"];
	const WorkbookView = AscCommonExcel.WorkbookView;
	const api = window["Asc"]["spreadsheet_api"];

	const MAX_CELLS_SCANNED = 100000;

	function colName(c) {
		let s = "";
		c++;
		while (c > 0) {
			const m = (c - 1) % 26;
			s = String.fromCharCode(65 + m) + s;
			c = Math.floor((c - 1) / 26);
		}
		return s;
	}

	/**
	 * @returns {?{x: number, y: number, w: number, h: number, numeric: boolean, rows: number, cols: number}}
	 * the bottom-right cell of the selection (pixels, relative to the grid), when Quick Analysis applies
	 */
	function getInfo(wb) {
		if (!wb || wb.getCellEditMode() || !wb.Api.canEdit()) {
			return null;
		}
		const ws = wb.getWorksheet();
		if (!ws || !ws.model || ws.getSelectionShape && ws.getSelectionShape()) {
			return null;
		}
		const render = ws.objectRender;
		if (render && render.getSelectedGraphicObjects && render.getSelectedGraphicObjects().length) {
			return null;
		}
		const sr = ws.model.selectionRange;
		if (!sr || sr.ranges.length !== 1) {
			return null;
		}
		const r = sr.getLast();
		if (r.isOneCell() || r.getType() !== Asc.c_oAscSelectionType.RangeCells) {
			return null;
		}
		let hasData = false, numeric = false, scanned = 0;
		ws.model.getRange3(r.r1, r.c1, r.r2, r.c2)._foreachNoEmpty(function (cell) {
			if (++scanned > MAX_CELLS_SCANNED || (hasData && numeric)) {
				return true;
			}
			if (!cell.isNullText()) {
				hasData = true;
				if (cell.getType() === AscCommon.CellValueType.Number) {
					numeric = true;
				}
			}
		});
		if (!hasData) {
			return null;
		}
		const vr = ws.visibleRange;
		if (r.r2 < vr.r1 || r.c2 < vr.c1 || r.r2 > vr.r2 + 1 || r.c2 > vr.c2 + 1) {
			return null;
		}
		const rect = ws.getCellCoord(r.c2, r.r2);
		return {
			x: rect.asc_getX(), y: rect.asc_getY(), w: rect.asc_getWidth(), h: rect.asc_getHeight(),
			numeric: numeric, rows: r.r2 - r.r1 + 1, cols: r.c2 - r.c1 + 1
		};
	}

	function notify(wb) {
		if (wb.elixirQuickAnalysisTimer) {
			return;
		}
		wb.elixirQuickAnalysisTimer = setTimeout(function () {
			wb.elixirQuickAnalysisTimer = null;
			const info = getInfo(wb);
			const key = info ? [info.x, info.y, info.w, info.h, info.numeric].join() : "";
			if (key !== wb.elixirQuickAnalysisKey) {
				wb.elixirQuickAnalysisKey = key;
				wb.Api.sendEvent("asc_onQuickAnalysis", info);
			}
		}, 0);
	}

	["_onWSSelectionChanged", "_onScrollY", "_onScrollX", "changeZoom", "_onEditCell", "closeCellEditor", "showWorksheet"].forEach(function (name) {
		const original = WorkbookView.prototype[name];
		if (!original) {
			return;
		}
		WorkbookView.prototype[name] = function () {
			const res = original.apply(this, arguments);
			notify(this);
			return res;
		};
	});

	api.prototype.asc_getQuickAnalysisInfo = function () {
		return getInfo(this.wb);
	};

	/**
	 * Quick Analysis > Totals: a formula for each column in the row below the selection
	 * (byRow false), or for each row in the column to its right (byRow true).
	 * @param func {string} SUM, AVERAGE, COUNT, MAX or MIN
	 * @returns {boolean} false when the target row or column is not empty (nothing is written)
	 */
	api.prototype.asc_QuickAnalysisTotals = function (func, byRow) {
		if (!/^(SUM|AVERAGE|COUNT|MAX|MIN)$/.test(func) || this.collaborativeEditing.getGlobalLock() || !this.canEdit() || !this.wb) {
			return false;
		}
		const wb = this.wb, ws = wb.getWorksheet(), model = ws.model;
		const r = model.selectionRange.getLast().clone();
		const target = byRow ? new Asc.Range(r.c2 + 1, r.r1, r.c2 + 1, r.r2) : new Asc.Range(r.c1, r.r2 + 1, r.c2, r.r2 + 1);
		if (target.r2 > AscCommon.gc_nMaxRow0 || target.c2 > AscCommon.gc_nMaxCol0) {
			return false;
		}
		// never overwrite data: Excel's totals go into an empty row or column
		let occupied = false;
		model.getRange3(target.r1, target.c1, target.r2, target.c2)._foreachNoEmpty(function (cell) {
			if (!cell.isNullText()) {
				occupied = true;
				return true;
			}
		});
		if (occupied) {
			return false;
		}
		const localName = AscCommonExcel.cFormulaFunctionToLocale ? AscCommonExcel.cFormulaFunctionToLocale[func] || func : func;
		ws._isLockedCells(target, null, function (success) {
			if (!success) {
				return;
			}
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			if (byRow) {
				for (let row = r.r1; row <= r.r2; ++row) {
					model.getRange3(row, target.c1, row, target.c1).setValue("=" + localName + "(" + colName(r.c1) + (row + 1) + ":" + colName(r.c2) + (row + 1) + ")");
				}
			} else {
				for (let col = r.c1; col <= r.c2; ++col) {
					model.getRange3(target.r1, col, target.r1, col).setValue("=" + localName + "(" + colName(col) + (r.r1 + 1) + ":" + colName(col) + (r.r2 + 1) + ")");
				}
			}
			AscCommon.History.EndTransaction();
			ws._updateRange(target);
			ws.draw();
		});
		return true;
	};
	api.prototype["asc_getQuickAnalysisInfo"] = api.prototype.asc_getQuickAnalysisInfo;
	api.prototype["asc_QuickAnalysisTotals"] = api.prototype.asc_QuickAnalysisTotals;
})(window);
