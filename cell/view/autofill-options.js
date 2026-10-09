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
 * Auto Fill Options, as in Microsoft Excel: after a fill handle drag (or a double click on the handle)
 * the interface shows a button at the bottom-right corner of the filled range (event
 * asc_onAutoFillOptions with its position and the option in use, or null) whose menu offers Copy Cells,
 * Fill Series, Fill Formatting Only, Fill Without Formatting and Flash Fill. Choosing an option undoes
 * the fill and does it again the new way, like Excel: the history keeps one step.
 */
(function (window) {
	const AscCommonExcel = window["AscCommonExcel"];
	const WorkbookView = AscCommonExcel.WorkbookView;
	const WorksheetView = AscCommonExcel.WorksheetView;
	const History = AscCommon.History;
	const api = window["Asc"]["spreadsheet_api"];
	const FT = Asc.c_oAscFillType;

	const MAX_CELLS = 200000;

	function sameRange(a, b) {
		return a && b && a.c1 === b.c1 && a.r1 === b.r1 && a.c2 === b.c2 && a.r2 === b.r2;
	}

	/** the cells added by the fill: the filled range minus the source */
	function targetOf(ctx, after) {
		const s = ctx.source;
		if (ctx.dir === 1) {
			return after.r2 > s.r2 ? new Asc.Range(s.c1, s.r2 + 1, s.c2, after.r2) : new Asc.Range(s.c1, after.r1, s.c2, s.r1 - 1);
		}
		return after.c2 > s.c2 ? new Asc.Range(s.c2 + 1, s.r1, after.c2, s.r2) : new Asc.Range(after.c1, s.r1, s.c1 - 1, s.r2);
	}

	function isEligible(ws) {
		return ws.activeFillHandle && null === ws.resizeTableIndex && 2 !== ws.fillHandleArea &&
			!ws.model.getActiveFillType() && !ws.model.getFillHandleRightClick() && ws.model.selectionRange.ranges.length === 1;
	}

	/** which option the fill used: one numeric cell is copied by a plain drag and extended as a series with Ctrl */
	function defaultOption(ws, ctx) {
		const s = ctx.source;
		let copy = ctx.ctrl;
		if (s.isOneCell()) {
			const range = ws.model.getRange3(s.r1, s.c1, s.r2, s.c2);
			const fmt = range.getXfs() && range.getXfs().num && range.getXfs().num.getFormat();
			const isDate = !!(fmt && AscCommon.oNumFormatCache.get(fmt).isDateTimeFormat());
			if (range.getType() === AscCommon.CellValueType.Number && !isDate) {
				copy = !copy;
			}
		}
		return copy ? FT.copyCells : FT.fillSeries;
	}

	function record(ws, ctx, selected) {
		const after = ws.model.selectionRange.getLast().clone();
		ws.elixirAutoFill = {ctx: ctx, selected: selected, after: after};
		notify(ws.getApi().wb);
	}

	const originalApply = WorksheetView.prototype.applyFillHandle;
	WorksheetView.prototype.applyFillHandle = function (x, y, ctrlPress, opt_doNotDraw, callback) {
		const ws = this;
		const replay = ws.elixirAutoFillReplay;
		ws.elixirAutoFillReplay = null;
		let ctx = null;
		if (replay) {
			ctx = replay.ctx;
		} else if (isEligible(ws)) {
			ctx = {
				source: ws.model.selectionRange.getLast().clone(), handle: ws.activeFillHandle.clone(),
				dir: ws.fillHandleDirection, area: ws.fillHandleArea, ctrl: !!ctrlPress
			};
		}
		ws.elixirAutoFill = null;
		return originalApply.call(this, x, y, ctrlPress, opt_doNotDraw, function (ok) {
			callback && callback(ok);
			if (replay && replay.after) {
				replay.after(ok);
			}
			if (ok && ctx) {
				record(ws, ctx, replay ? replay.type : defaultOption(ws, ctx));
			} else if (!ok && replay) {
				ws.getApi().wb && notify(ws.getApi().wb);
			}
		});
	};

	/**
	 * @returns {?{x: number, y: number, w: number, h: number, selected: number, flash: boolean}}
	 * the bottom-right cell of the filled range (pixels, relative to the grid)
	 */
	function getInfo(wb) {
		if (!wb || wb.getCellEditMode() || !wb.Api.canEdit()) {
			return null;
		}
		const ws = wb.getWorksheet();
		const af = ws && ws.elixirAutoFill;
		if (!af) {
			return null;
		}
		const sr = ws.model.selectionRange;
		if (!sr || sr.ranges.length !== 1 || !sameRange(sr.getLast(), af.after)) {
			ws.elixirAutoFill = null;
			return null;
		}
		const r = af.after, vr = ws.visibleRange;
		if (r.r2 < vr.r1 || r.c2 < vr.c1 || r.r2 > vr.r2 + 1 || r.c2 > vr.c2 + 1) {
			return null;
		}
		const rect = ws.getCellCoord(r.c2, r.r2);
		return {
			x: rect.asc_getX(), y: rect.asc_getY(), w: rect.asc_getWidth(), h: rect.asc_getHeight(),
			selected: af.selected, flash: af.ctx.dir === 1 && af.ctx.source.c1 === af.ctx.source.c2
		};
	}

	function notify(wb) {
		if (wb.elixirAutoFillTimer) {
			return;
		}
		wb.elixirAutoFillTimer = setTimeout(function () {
			wb.elixirAutoFillTimer = null;
			const info = getInfo(wb);
			const key = info ? [info.x, info.y, info.w, info.h, info.selected].join() : "";
			if (key !== wb.elixirAutoFillKey) {
				wb.elixirAutoFillKey = key;
				wb.Api.sendEvent("asc_onAutoFillOptions", info);
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

	/** true while the Auto Fill Options button belongs to the current selection (Quick Analysis then stays hidden) */
	AscCommonExcel.autoFillOptionsActive = function (ws) {
		const af = ws.elixirAutoFill;
		if (!af) {
			return false;
		}
		const sr = ws.model.selectionRange;
		if (!sr || sr.ranges.length !== 1 || !sameRange(sr.getLast(), af.after)) {
			ws.elixirAutoFill = null;
			return false;
		}
		return true;
	};

	api.prototype.asc_getAutoFillOptionsInfo = function () {
		return getInfo(this.wb);
	};

	/**
	 * Re-does the last fill with another option.
	 * @param type {Asc.c_oAscFillType} copyCells, fillSeries, fillFormattingOnly, fillWithoutFormatting or flashFill
	 * @returns {boolean} false when there is no fill to change or the option does not apply
	 */
	api.prototype.asc_ApplyAutoFillOption = function (type) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit() || !this.wb || this.wb.getCellEditMode()) {
			return false;
		}
		const wb = this.wb, ws = wb.getWorksheet(), af = ws.elixirAutoFill;
		if (!af || !AscCommonExcel.autoFillOptionsActive(ws)) {
			return false;
		}
		const ctx = af.ctx, previous = af.selected;
		if (type === af.selected) {
			return true;
		}
		const isFlash = type === FT.flashFill;
		if (isFlash && !(ctx.dir === 1 && ctx.source.c1 === ctx.source.c2)) {
			return false;
		}
		const target = targetOf(ctx, af.after);
		if ((target.r2 - target.r1 + 1) * (target.c2 - target.c1 + 1) > MAX_CELLS) {
			return false;
		}
		if (ws.model.getSheetProtection() || ws.model.isUserProtectedRangesIntersection(target)) {
			return false;
		}
		ws.elixirAutoFill = null;
		this.asc_Undo();
		ws.setSelection(ctx.source.clone());
		const restoreHandle = function () {
			ws.activeFillHandle = ctx.handle.clone();
			ws.fillHandleDirection = ctx.dir;
			ws.fillHandleArea = ctx.area;
		};
		if (isFlash) {
			const res = this.asc_FlashFill();
			if (res < 0) {
				restoreHandle();
				ws.elixirAutoFillReplay = {ctx: ctx, type: previous};
				ws.applySeriesSettings(previous);
				return false;
			}
			ws.setSelection(af.after.clone());
			record(ws, ctx, FT.flashFill);
			return true;
		}
		let kept = null;
		if (type === FT.fillWithoutFormatting) {
			// the cells keep the formatting they had before the fill
			kept = [];
			for (let r = target.r1; r <= target.r2; ++r) {
				for (let c = target.c1; c <= target.c2; ++c) {
					kept.push([r, c, ws.model.getRange3(r, c, r, c).getXfs()]);
				}
			}
		}
		const base = type === FT.fillFormattingOnly ? FT.copyCells : (type === FT.fillWithoutFormatting ? FT.fillSeries : type);
		History.Create_NewPoint();
		History.StartTransaction();
		restoreHandle();
		ws.elixirAutoFillReplay = {
			ctx: ctx, type: type, after: function (ok) {
				if (ok && type === FT.fillFormattingOnly) {
					ws.model.getRange3(target.r1, target.c1, target.r2, target.c2).cleanText();
				} else if (ok && kept) {
					kept.forEach(function (k) {
						ws.model.getRange3(k[0], k[1], k[0], k[1]).setStyle(k[2]);
					});
				}
				History.EndTransaction();
				ws._updateRange(af.after);
				ws.draw();
			}
		};
		ws.applySeriesSettings(base);
		return true;
	};
	api.prototype["asc_getAutoFillOptionsInfo"] = api.prototype.asc_getAutoFillOptionsInfo;
	api.prototype["asc_ApplyAutoFillOption"] = api.prototype.asc_ApplyAutoFillOption;
})(window);
