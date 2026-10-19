/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
"use strict";
(function(window)
{
	const Direction = {};
	Direction.parentMovesRight = function(parent, isRight)
	{
		let document = parent.LogicDocument || parent;
		let move = document._tableCursorMove;
		if (!move || !parent.Content || !parent.CurPos)
			return isRight;
		// Preserve document-order collapse/expansion across multiple elements.
		if (parent.Selection.Use && parent.Selection.StartPos !== parent.Selection.EndPos)
			return isRight;
		let index = parent.Selection.Use ? parent.Selection.EndPos : parent.CurPos.ContentPos;
		let child = parent.Content[index];
		if (!child || !child.IsTable || !child.IsTable())
			return isRight;
		return move.physicalRight !== child.IsRTL();
	};
	Direction.reflectRulerMarkup = function(markup, axis)
	{
		let width = markup.Cols.reduce(function(sum, column) { return sum + column; }, 0);
		markup.X = axis - markup.X - width;
		markup.Cols.reverse();
		markup.Margins = markup.Margins.slice().reverse().map(function(margin)
		{
			return {Left : margin.Right, Right : margin.Left};
		});
		markup.CurCol = markup.Cols.length - 1 - markup.CurCol;
	};
	Direction.moveRtlRulerBoundary = function(markup, index, position, pageWidth)
	{
		let boundary = markup.X;
		for (let i = 0; i < index; ++i)
			boundary += markup.Cols[i];
		let maximum = pageWidth;
		if (index < markup.Cols.length)
			maximum = boundary + markup.Cols[index] - markup.Margins[index].Left - markup.Margins[index].Right;
		position = Math.max(0, Math.min(maximum, position));
		let delta = position - boundary;
		markup.X += delta;
		if (index < markup.Cols.length)
			markup.Cols[index] -= delta;
		return position;
	};
	Direction.resizeHandleBounds = function(bounds, rtl, size)
	{
		let left = rtl ? bounds.X - size : bounds.X + bounds.W;
		let top = bounds.Y + bounds.H;
		return {Left : left, Right : left + size, Top : top, Bottom : top + size};
	};
	Direction.resizePointerDelta = function(start, point, rtl)
	{
		return {W : (rtl ? -1 : 1) * (point.X - start.X), H : point.Y - start.Y};
	};
	Direction.reflectMetrics = function(info, width)
	{
		const result = Object.assign({}, info);
		for (const kind of ["grid", "cell", "content"])
		{
			result["X_" + kind + "_start"] = width - info["X_" + kind + "_end"];
			result["X_" + kind + "_end"] = width - info["X_" + kind + "_start"];
		}
		return result;
	};
	Direction.reflectTablePosition = function(x, left, right, width)
	{
		return left + right - width - x;
	};
	Direction.physicalBorders = function(borders, rtl)
	{
		if (!rtl)
			return borders;
		return Object.assign({}, borders, {Left: borders.Right, Right: borders.Left});
	};
	Direction.createBorderGraphics = function(graphics, pageX, width)
	{
		// Only table border strokes use this adapter. Content uses its already
		// positioned physical metrics and never receives reflected graphics.
		const axis = 2 * pageX + width;
		const rule = window.AscCommon.c_oAscLineDrawingRule;
		const overrides = {
			drawVerLine: function(align, x, top, bottom, penWidth)
			{
				if (align === rule.Left) align = rule.Right;
				else if (align === rule.Right) align = rule.Left;
				return graphics.drawVerLine(align, axis - x, top, bottom, penWidth);
			},
			drawHorLineExt: function(align, y, left, right, penWidth, leftExtension, rightExtension)
			{
				return graphics.drawHorLineExt(align, y, axis - right, axis - left, penWidth, -rightExtension, -leftExtension);
			},
			DrawEmptyTableLine: function(x1, y1, x2, y2)
			{
				return graphics.DrawEmptyTableLine(axis - x1, y1, axis - x2, y2);
			}
		};
		const bound = new Map();
		return new Proxy(graphics, {
			get: function(target, key)
			{
				if (Object.prototype.hasOwnProperty.call(overrides, key))
					return overrides[key];
				const value = Reflect.get(target, key, target);
				if (typeof value !== "function")
					return value;
				let cached = bound.get(key);
				if (!cached || cached.source !== value)
				{
					cached = {source: value, method: value.bind(target)};
					bound.set(key, cached);
				}
				return cached.method;
			}
		});
	};
	window["AscWord"].TableDirection = window.AscWord.TableDirection = Direction;
})(window);
