/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";
(function(window)
{
	// Run against a loaded real editor and an RTL table fixture.
	// Drawing checks use a recording overlay with the actual page/matrix methods.
	window["AscTestTableDirection"] = window.AscTestTableDirection = function()
	{
		const api = window.Asc.editor;
		const doc = api.WordControl.m_oLogicDocument;
		const table = doc.Content.find(element => element.IsTable && element.IsTable());
		const results = [];
		function equal(actual, expected, message)
		{
			if (actual !== expected)
				throw new Error(message + ": " + String(actual) + " != " + String(expected));
		}
		function check(name, callback)
		{
			try { callback(); results.push({name: name, status: "PASS"}); }
			catch (error) { results.push({name: name, status: "FAIL", error: String(error)}); }
		}
		function action(callback)
		{
			doc.StartAction(AscDFH.historydescription_Document_ApplyTablePr);
			try { callback(); doc.Recalculate(); }
			finally { doc.FinalizeAction(); }
		}
		const TablePr = table.Pr.constructor;
		check("actual memory/stream internal property three-state boundary", function()
		{
			for (const direction of [true, false, undefined])
			{
				const pr = new TablePr();
				pr.BidiVisual = direction;
				pr.TableCaption = "boundary";
				const memory = new AscCommon.CMemory();
				pr.Write_ToBinary(memory);
				const boundary = memory.GetCurPosition();
				memory.WriteLong(0x12345678);
				const data = memory.GetData();
				const stream = new AscCommon.FT_Stream2(data, data.length);
				const copy = new TablePr();
				copy.Read_FromBinary(stream);
				equal(copy.BidiVisual, direction, "internal direction");
				equal(copy.TableCaption, "boundary", "caption");
				equal(stream.cur, boundary, "record boundary");
				equal(stream.GetLong(), 0x12345678, "following record");
			}
		});
		check("real bool history serialization, undo and redo for all nine state pairs", function()
		{
			const Change = AscDFH.changesFactory[AscDFH.historyitem_Table_BidiVisual];
			const original = table.Pr.BidiVisual;
			try
			{
				for (const oldValue of [true, false, undefined])
				{
					for (const newValue of [true, false, undefined])
					{
						const memory = new AscCommon.CMemory();
						new Change(table, oldValue, newValue).WriteToBinary(memory);
						const data = memory.GetData();
						const copy = new Change(table);
						copy.ReadFromBinary(new AscCommon.FT_Stream2(data, data.length));
						equal(copy.Old, oldValue, "history old");
						equal(copy.New, newValue, "history new");
						copy.Redo();
						equal(table.Pr.BidiVisual, newValue, "redo");
						copy.Undo();
						equal(table.Pr.BidiVisual, oldValue, "undo");
					}
				}
			}
			finally { table.Pr.BidiVisual = original; table.Recalc_CompiledPr2(); }
		});
		check("RTL pencil and eraser remain gated", function()
		{
			equal(table.IsRTL(), true, "RTL fixture");
			const before = JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true),
				rows: table.Content.map(row => row.Content.length)});
			table.DrawTableCells(0, 0, 100, 100, 0, 0, true);
			table.DrawTableCells(0, 0, 100, 100, 0, 0, false);
			equal(JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true),
				rows: table.Content.map(row => row.Content.length)}), before, "pencil/eraser preserve table");
		});
		check("RTL physical borders expose logical drag indices", function()
		{
			const cell = table.Content[0].Content[1];
			const y = table.RowsInfo[0].Y[0] + table.RowsInfo[0].H[0] / 2;
			const left = table.private_CheckHitInBorder(table.Pages[0].X + cell.Metrics.X_grid_start, y, 0);
			const right = table.private_CheckHitInBorder(table.Pages[0].X + cell.Metrics.X_grid_end, y, 0);
			equal(left.Border === 1 || left.Border === 3, true, "physical left is a column border");
			equal(right.Border === 1 || right.Border === 3, true, "physical right is a column border");
			equal(left.Pos.Cell + (left.Border === 1 ? 1 : 0), 2, "physical left maps to logical boundary two");
			equal(right.Pos.Cell + (right.Border === 1 ? 1 : 0), 1, "physical right maps to logical boundary one");
		});
		check("RTL outline resize mutates width, keeps the right edge and restores history", function()
		{
			const before = JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)});
			doc.RemoveSelection();
			table.CurCell.Content.Set_CurrentElement(0, true);
			const bounds = table.GetPageBounds(table.Pages.length - 1);
			const right = table.GetPageBounds(0).Right;
			table.ResizeTableInDocument(bounds.Right - bounds.Left + 5, bounds.Bottom - bounds.Top + 2, true);
			const after = JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)});
			equal(after !== before, true, "outline resize changes properties");
			equal(Math.abs(table.GetPageBounds(0).Right - right) < 0.05, true, "physical right edge stays fixed");
			api.Undo();
			equal(JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)}), before, "resize undo");
			api.Redo();
			equal(JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)}), after, "resize redo");
			api.Undo();
		});
		check("RTL ruler mapping mutates a column and restores history", function()
		{
			table.private_UpdateTableMarkup(0, 0, 0);
			const markup = table.Markup.CreateDublicate();
			const axis = 2 * table.Pages[0].X + table.TableSumGrid[table.TableSumGrid.length - 1];
			AscWord.TableDirection.reflectRulerMarkup(markup, axis);
			markup.TableDirectionRTL = true;
			markup.TableDirectionAxis = axis;
			markup.TableDirectionGeometry = table.GetInteractionGeometry();
			const before = JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)});
			AscWord.TableDirection.moveRtlRulerBoundary(markup, 1, markup.X + markup.Cols[0] + 3, 1000);
			action(function() { table.Update_TableMarkupFromRuler(markup, true, 1); });
			const after = JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)});
			equal(after !== before, true, "ruler edit changes properties");
			api.Undo();
			equal(JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)}), before, "ruler undo");
			api.Redo();
			equal(JSON.stringify({grid: table.TableGrid, pr: table.Pr.Copy(true)}), after, "ruler redo");
			api.Undo();
		});
		check("actual transformed outline paint matches its RTL local hit rectangle", function()
		{
			const drawing = api.WordControl.m_oDrawingDocument;
			const tracker = drawing.TableOutlineDr;
			const page = drawing.m_arrPages[0];
			const original = tracker.TableMatrix;
			const originalOutline = tracker.TableOutline;
			const bounds = {X : 30, Y : 50, W : 80, H : 20};
			const matrix = new AscCommon.CMatrix();
			const angle = Math.PI / 6;
			matrix.sx = matrix.sy = 2 * Math.cos(angle);
			matrix.shx = -2 * Math.sin(angle);
			matrix.shy = 2 * Math.sin(angle);
			matrix.tx = 10; matrix.ty = 20;
			const points = [];
			const overlay = {CheckPoint : function(x, y) { points.push({X : x, Y : y}); },
				m_oContext : {beginPath : function() {}, moveTo : function() {}, lineTo : function() {},
					closePath : function() {}, fill : function() {}, stroke : function() {}}};
			try
			{
				tracker.TableMatrix = matrix;
				tracker.TableOutline = {Table : table};
				const scale = 96 / 25.4 * api.WordControl.m_nZoomValue / 100;
				page.DrawTableOutline(overlay, 0, 0, page.width_mm * scale, page.height_mm * scale, tracker, bounds);
				equal(points.length, 4, "painted quad");
				const dpr = AscCommon.AscBrowser.retinaPixelRatio;
				const coords = new AscCommon.CMatrix();
				coords.sx = coords.sy = scale * dpr;
				const composed = matrix.CreateDublicate();
				AscCommon.global_MatrixTransformer.MultiplyAppend(composed, coords);
				const inverse = AscCommon.global_MatrixTransformer.Invert(composed);
				const local = points.map(function(point)
				{
					return {X : inverse.TransformPointX(point.X, point.Y), Y : inverse.TransformPointY(point.X, point.Y)};
				});
				const hit = AscWord.TableDirection.resizeHandleBounds(bounds, true, tracker.getResizeHandleSize(api.WordControl));
				equal(Math.abs(Math.min(...local.map(point => point.X)) - hit.Left) < 0.0001, true, "painted left matches hit");
				equal(Math.abs(Math.max(...local.map(point => point.X)) - hit.Right) < 0.0001, true, "painted right matches hit");
				equal(Math.abs(Math.max(...local.map(point => point.Y)) - hit.Bottom) < 0.0001, true, "painted bottom matches hit");
			}
			finally { tracker.TableMatrix = original; tracker.TableOutline = originalOutline; }
		});
		check("editor history action undo/redo", function()
		{
			equal(table.IsRTL(), true, "input is RTL");
			const starts = function() { return table.Content[0].Content.map(cell => cell.Metrics.X_cell_start); };
			equal(starts()[0] > starts()[2], true, "initial RTL geometry");
			table.CurCell.Content.MoveCursorToStartPos();
			table.CurCell.Content.Set_CurrentElement(0, true);
			const properties = new Asc.CTableProp();
			equal(properties.get_BidiVisual(), undefined, "unspecified API direction");
			properties.put_BidiVisual(false);
			api.tblApply(properties);
			equal(new Asc.CTableProp(table.Get_Props()).get_BidiVisual(), false, "reported API direction");
			equal(table.IsRTL(), false, "changed direction");
			equal(starts()[0] < starts()[2], true, "changed LTR geometry");
			api.Undo();
			equal(table.IsRTL(), true, "editor undo");
			equal(starts()[0] > starts()[2], true, "undo RTL geometry");
			api.Redo();
			equal(table.IsRTL(), false, "editor redo");
			equal(starts()[0] < starts()[2], true, "redo LTR geometry");
		});
		check("full-property change absorbs an explicit false direction", function()
		{
			const Change = AscDFH.changesFactory[AscDFH.historyitem_Table_BidiVisual];
			const Full = AscDFH.changesFactory[AscDFH.historyitem_Table_Pr];
			const before = table.Pr.Copy(true), after = table.Pr.Copy(true);
			before.BidiVisual = true; after.BidiVisual = true;
			const full = new Full(table, before, after);
			full.Merge(new Change(table, true, false));
			equal(full.New.BidiVisual, false, "merged new direction");
			equal(full.Old.BidiVisual, true, "unchanged old direction");
		});
		check("compiled style inheritance and explicit false override", function()
		{
			action(function()
			{
				const style = new AscCommonWord.CStyle("RTL direction test", null, null, AscWord.styletype_Table);
				style.TablePr.BidiVisual = true;
				const id = table.Get_Styles().Add(style);
				table.SetTableStyle(id);
				table.Set_BidiVisual(undefined);
				equal(table.IsRTL(), true, "inherited true");
				table.Set_BidiVisual(false);
				equal(table.IsRTL(), false, "direct false overrides style");
				table.Set_TableStyle(id);
				equal(table.Pr.BidiVisual, undefined, "clear-direct style path clears direction");
				equal(table.IsRTL(), true, "style now controls direction");
				table.Set_BidiVisual(false);
				const other = new AscCommonWord.CStyle("Other RTL direction test", null, null, AscWord.styletype_Table);
				other.TablePr.BidiVisual = true;
				table.SetTableStyle(table.Get_Styles().Add(other));
				equal(table.Pr.BidiVisual, false, "no-clear style path retains direction");
				equal(table.IsRTL(), false, "explicit false remains effective");
			});
		});
		return {status: results.every(result => result.status === "PASS") ? "PASS" : "FAIL", results: results};
	};
})(window);
