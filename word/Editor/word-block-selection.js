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
 * Block (column) selection of Word:
 *   Alt+drag (macOS Option+drag)           select a vertical block of text
 *   Ctrl+Shift+F8 (macOS Cmd+Shift+F8)     start a block at the cursor, then the arrow keys extend it; Esc ends it
 * The block is a rectangle on one page; on each line of body text it crosses, the characters between its left and
 * right edges are selected. Copy and cut give one line of text per line of the block; Delete and Backspace remove
 * the characters of the block; character formatting (bold, font, size, colour...) applies to the block only.
 * Any other key or click ends the block (typing starts at its top left corner).
 *
 * The engine keeps an ordinary selection from the first to the last character of the block (so that the
 * interface shows the formatting of the selection), but draws only the block.
 */
(function(window)
{
	const PreventAll = 0xFFFF;
	const KEY_STEP   = 0.5;   // mm: a key press moves the edge of the block to the next character boundary

	// ---------------------------------------------------------------- paragraph offsets
	function countItems(item)
	{
		if (item instanceof ParaRun)
			return item.Content.length;
		if (item && Array.isArray(item.Content))
		{
			let n = 0;
			for (let i = 0; i < item.Content.length; ++i)
				n += countItems(item.Content[i]);
			return n;
		}
		return 0;
	}

	function posToOffset(paragraph, pos)
	{
		let offset = 0;
		(function rec(container, depth)
		{
			let index = pos.Get(depth);
			for (let i = 0; i < container.Content.length && i <= index; ++i)
			{
				let item = container.Content[i];
				if (i < index)
					offset += countItems(item);
				else if (item instanceof ParaRun)
					offset += Math.min(pos.Get(depth + 1) || 0, item.Content.length);
				else if (item && Array.isArray(item.Content))
					rec(item, depth + 1);
			}
		})(paragraph, 0);
		return offset;
	}

	function offsetToPos(paragraph, offset)
	{
		let pos  = new AscWord.CParagraphContentPos();
		let left = offset;
		let done = (function rec(container, depth)
		{
			for (let i = 0; i < container.Content.length; ++i)
			{
				let item = container.Content[i];
				let n = countItems(item);
				if (item instanceof ParaRun)
				{
					if (left <= n)
					{
						pos.Update(i, depth);
						pos.Update(left, depth + 1);
						return true;
					}
					left -= n;
				}
				else if (item && Array.isArray(item.Content) && n > 0)
				{
					if (left <= n)
					{
						pos.Update(i, depth);
						if (rec(item, depth + 1))
							return true;
					}
					else
						left -= n;
				}
			}
			return false;
		})(paragraph, 0);
		return done ? pos : paragraph.Get_EndPos(false);
	}

	// ---------------------------------------------------------------- block geometry
	function getBlock(doc)
	{
		return doc.ElxBlock && doc.ElxBlock.active ? doc.ElxBlock : null;
	}

	/**
	 * The lines crossed by the rectangle and, on each one, the characters between its edges
	 */
	function computeRanges(doc, block)
	{
		let page = block.anchor.page;
		let x0 = Math.min(block.anchor.x, block.focus.x), x1 = Math.max(block.anchor.x, block.focus.x);
		let y0 = Math.min(block.anchor.y, block.focus.y), y1 = Math.max(block.anchor.y, block.focus.y);
		let ranges = [];
		for (let index = 0; index < doc.Content.length; ++index)
		{
			let paragraph = doc.Content[index];
			if (!paragraph.IsParagraph() || !paragraph.IsRecalculated())
				continue;
			for (let p = 0; p < paragraph.Pages.length; ++p)
			{
				if (paragraph.GetAbsolutePage(p) !== page)
					continue;
				let pg = paragraph.Pages[p];
				for (let line = pg.StartLine; line <= pg.EndLine; ++line)
				{
					let bounds = paragraph.GetLineBounds(line);
					if (bounds.Bottom - bounds.Top < 0.01 || bounds.Bottom <= y0 || bounds.Top > y1)
						continue;
					let s = paragraph.getSearchPosByXY(x0, line, p, true, false, true).getPos();
					let e = paragraph.getSearchPosByXY(x1, line, p, true, false, true).getPos();
					if (!s || !e)
						continue;
					let start = posToOffset(paragraph, s), end = posToOffset(paragraph, e);
					if (end < start)
					{
						let t = start; start = end; end = t;
						let tp = s; s = e; e = tp;
					}
					let xs = paragraph.Get_XYByContentPos(s).X, xe = paragraph.Get_XYByContentPos(e).X;
					ranges.push({paragraph : paragraph, index : index, start : start, end : end, x : Math.min(xs, xe), w : Math.abs(xe - xs), y : bounds.Top, h : bounds.Bottom - bounds.Top, page : page});
				}
			}
		}
		return ranges;
	}

	// an ordinary selection from the first to the last character of the block
	function selectSpan(doc, ranges)
	{
		doc.RemoveSelection();
		let nonEmpty = ranges.filter(function(r) { return r.end > r.start; });
		let first = nonEmpty[0] || ranges[0], last = nonEmpty[nonEmpty.length - 1] || ranges[ranges.length - 1];
		if (!first)
			return;
		let a = first.index, b = last.index;
		doc.SetDocPosType(docpostype_Content);
		doc.CurPos.ContentPos = b;
		doc.Selection.Use      = true;
		doc.Selection.Start    = false;
		doc.Selection.Flag     = selectionflag_Common;
		doc.Selection.StartPos = a;
		doc.Selection.EndPos   = b;
		let sp = offsetToPos(first.paragraph, first.start), ep = offsetToPos(last.paragraph, last.end);
		for (let i = a; i <= b; ++i)
		{
			let paragraph = doc.Content[i];
			if (!paragraph.IsParagraph())
			{
				paragraph.SelectAll(1);
				continue;
			}
			paragraph.Selection.Use   = true;
			paragraph.Selection.Start = false;
			let s = i === a ? sp : paragraph.Get_StartPos();
			let e = i === b ? ep : paragraph.Get_EndPos(false);
			paragraph.Set_SelectionContentPos(s, e);
		}
	}

	function selectRange(doc, range)
	{
		doc.RemoveSelection();
		let paragraph = range.paragraph;
		doc.SetDocPosType(docpostype_Content);
		doc.CurPos.ContentPos = range.index;
		paragraph.Set_ParaContentPos(offsetToPos(paragraph, range.start), true, -1, -1);
		doc.StartSelectionFromCurPos();
		paragraph.Set_SelectionContentPos(offsetToPos(paragraph, range.start), offsetToPos(paragraph, range.end));
	}

	// places the cursor at the top left corner of the block
	function collapse(doc, block)
	{
		let first = block.ranges[0];
		doc.ElxBlock = null;
		doc.RemoveSelection();
		if (first)
		{
			doc.SetDocPosType(docpostype_Content);
			doc.CurPos.ContentPos = first.index;
			first.paragraph.Set_ParaContentPos(offsetToPos(first.paragraph, first.start), true, -1, -1);
			first.paragraph.RecalculateCurPos && first.paragraph.RecalculateCurPos();
		}
		doc.UpdateSelection();
		doc.UpdateInterface();
		doc.sendEvent("asc_onWordBlockSelection", false);
	}

	CDocument.prototype.ElxUpdateBlock = function()
	{
		let block = getBlock(this);
		if (!block)
			return;
		block.ranges = computeRanges(this, block);
		selectSpan(this, block.ranges);
		this.UpdateSelection();
		this.UpdateInterface();
		this.UpdateRulers && this.UpdateRulers();
	};

	// after a change of the text, the block keeps its characters (offsets), not its rectangle
	function refreshRects(doc, block)
	{
		block.ranges.forEach(function(r)
		{
			let paragraph = r.paragraph;
			if (!paragraph.IsRecalculated())
				return;
			let s = paragraph.Get_XYByContentPos(offsetToPos(paragraph, r.start));
			let e = paragraph.Get_XYByContentPos(offsetToPos(paragraph, r.end));
			r.x = Math.min(s.X, e.X);
			r.w = Math.abs(e.X - s.X);
		});
		selectSpan(doc, block.ranges);
		doc.UpdateSelection();
		doc.UpdateInterface();
	}

	CDocument.prototype.StartBlockSelection = function(x, y, page, byKeys)
	{
		if (docpostype_Content !== this.GetDocPosType())
			return false;
		if (undefined === x)
		{
			let paragraph = this.GetCurrentParagraph();
			if (!paragraph || paragraph.GetParent() !== this)
				return false;
			let xy = paragraph.GetCurPosXY ? paragraph.GetCurPosXY() : {X : this.CurPos.RealX, Y : this.CurPos.RealY};
			let line = paragraph.CurPos.Line;
			let bounds = paragraph.GetLineBounds(line >= 0 ? line : 0);
			x = xy.X;
			y = (bounds.Top + bounds.Bottom) / 2;
			page = paragraph.GetAbsolutePage(paragraph.getPageByLine(line >= 0 ? line : 0));
		}
		this.ElxBlock = {active : true, byKeys : !!byKeys, dragging : !byKeys, anchor : {x : x, y : y, page : page}, focus : {x : x, y : y, page : page}, ranges : []};
		this.ElxUpdateBlock();
		this.sendEvent("asc_onWordBlockSelection", true);
		return true;
	};

	CDocument.prototype.EndBlockSelection = function()
	{
		let block = getBlock(this);
		block && collapse(this, block);
	};

	CDocument.prototype.IsBlockSelection = function()
	{
		return !!getBlock(this);
	};

	CDocument.prototype.GetBlockSelectionText = function()
	{
		let block = getBlock(this);
		if (!block)
			return null;
		let doc = this, lines = [];
		block.ranges.forEach(function(r)
		{
			selectRange(doc, r);
			lines.push(r.end > r.start ? (doc.GetSelectedText(false) || "").replace(/[\r\n]+$/, "") : "");
		});
		selectSpan(this, block.ranges);
		return lines;
	};

	// Delete, Backspace, Cut: the characters of the block, in one undo step
	CDocument.prototype.DeleteBlockSelection = function()
	{
		let block = getBlock(this);
		if (!block)
			return false;
		let ranges = block.ranges.filter(function(r) { return r.end > r.start; });
		if (!ranges.length)
		{
			collapse(this, block);
			return true;
		}
		selectSpan(this, block.ranges);
		if (this.IsSelectionLocked(AscCommon.changestype_Remove, null, true, this.IsFormFieldEditing()))
			return true;
		this.StartAction(AscDFH.historydescription_Document_DeleteButton);
		for (let i = ranges.length - 1; i >= 0; --i)
		{
			selectRange(this, ranges[i]);
			this.Remove(1, false, true, false);
		}
		this.Recalculate();
		this.FinalizeAction();
		collapse(this, block);
		return true;
	};

	// character formatting applies to each line of the block
	let addToParagraph = CDocument.prototype.AddToParagraph;
	CDocument.prototype.AddToParagraph = function(item, isRecalculate)
	{
		let block = getBlock(this);
		if (!block || !item || para_TextPr !== item.Type || block.applying)
			return addToParagraph.apply(this, arguments);
		let doc = this;
		block.applying = true;
		try
		{
			block.ranges.forEach(function(r)
			{
				if (r.end <= r.start)
					return;
				selectRange(doc, r);
				addToParagraph.call(doc, item, false);
			});
		}
		finally
		{
			block.applying = false;
		}
		this.Recalculate();
		refreshRects(this, block);
	};

	let drawSelectionOnPage = CDocument.prototype.DrawSelectionOnPage;
	CDocument.prototype.DrawSelectionOnPage = function(pageAbs)
	{
		let block = getBlock(this);
		if (!block)
			return drawSelectionOnPage.apply(this, arguments);
		this.DrawingDocument.UpdateTargetTransform(null);
		this.DrawingDocument.SetTextSelectionOutline(false);
		let dd = this.DrawingDocument;
		block.ranges.forEach(function(r)
		{
			if (r.page === pageAbs && r.w > 0.01)
				dd.AddPageSelection(pageAbs, r.x, r.y, r.w, r.h);
		});
	};

	// ---------------------------------------------------------------- mouse
	let onMouseDown = CDocument.prototype.OnMouseDown;
	CDocument.prototype.OnMouseDown = function(e, x, y, pageIndex)
	{
		let block = getBlock(this);
		if (e && e.AltKey && !e.CtrlKey && !e.ShiftKey && !e.MacCmdKey && 1 === e.ClickCount && e.Button !== 2 && docpostype_Content === this.GetDocPosType() && !this.IsViewMode() && !this.IsFillingFormMode())
		{
			let inDrawing = this.DrawingObjects && this.DrawingObjects.IsInDrawingObject ? this.DrawingObjects.IsInDrawingObject(x, y, pageIndex, this) : -1;
			if (inDrawing < 0 || inDrawing === DRAWING_ARRAY_TYPE_BEHIND)
			{
				this.CurPage = pageIndex;
				this.StartBlockSelection(x, y, pageIndex, false);
				return;
			}
		}
		if (block)
			this.ElxBlock = null;
		return onMouseDown.apply(this, arguments);
	};

	let onMouseMove = CDocument.prototype.OnMouseMove;
	CDocument.prototype.OnMouseMove = function(e, x, y, pageIndex)
	{
		let block = getBlock(this);
		if (block && block.dragging)
		{
			if (pageIndex === block.anchor.page)
			{
				block.focus = {x : x, y : y, page : pageIndex};
				this.ElxUpdateBlock();
			}
			return;
		}
		return onMouseMove.apply(this, arguments);
	};

	let onMouseUp = CDocument.prototype.OnMouseUp;
	CDocument.prototype.OnMouseUp = function(e, x, y, pageIndex)
	{
		let block = getBlock(this);
		if (block && block.dragging)
		{
			if (pageIndex === block.anchor.page)
				block.focus = {x : x, y : y, page : pageIndex};
			block.dragging = false;
			this.ElxUpdateBlock();
			if (!block.ranges.some(function(r) { return r.end > r.start; }))
				collapse(this, block);
			return;
		}
		return onMouseUp.apply(this, arguments);
	};

	// ---------------------------------------------------------------- keys
	function moveFocusX(doc, block, dir)
	{
		let line = lineAt(doc, block.focus);
		if (!line)
			return;
		let paragraph = line.paragraph;
		let at = function(x) { return posToOffset(paragraph, paragraph.getSearchPosByXY(x, line.line, line.page, true, false, true).getPos()); };
		let from = at(block.focus.x), x = block.focus.x;
		for (let i = 0; i < 200; ++i)
		{
			x += dir * KEY_STEP;
			if (x < 0)
				break;
			let off = at(x);
			if (off !== from)
			{
				block.focus.x = paragraph.Get_XYByContentPos(offsetToPos(paragraph, off)).X;
				return;
			}
		}
	}

	function lineAt(doc, point)
	{
		for (let index = 0; index < doc.Content.length; ++index)
		{
			let paragraph = doc.Content[index];
			if (!paragraph.IsParagraph())
				continue;
			for (let p = 0; p < paragraph.Pages.length; ++p)
			{
				if (paragraph.GetAbsolutePage(p) !== point.page)
					continue;
				let pg = paragraph.Pages[p];
				for (let line = pg.StartLine; line <= pg.EndLine; ++line)
				{
					let b = paragraph.GetLineBounds(line);
					if (point.y >= b.Top && point.y < b.Bottom)
						return {paragraph : paragraph, line : line, page : p, top : b.Top, bottom : b.Bottom, index : index};
				}
			}
		}
		return null;
	}

	function moveFocusY(doc, block, dir)
	{
		let line = lineAt(doc, block.focus);
		if (!line)
			return;
		let y = dir > 0 ? line.bottom + 0.05 : line.top - 0.05;
		for (let i = 0; i < 400; ++i, y += dir * 0.5)
		{
			let next = lineAt(doc, {x : block.focus.x, y : y, page : block.focus.page});
			if (next && next.bottom - next.top > 0.01)
			{
				block.focus.y = (next.top + next.bottom) / 2;
				return;
			}
			if (y < 0 || y > 1000)
				return;
		}
	}

	function isBlockStartKey(e)
	{
		let isMac = AscCommon.AscBrowser.isMacOs;
		// macOS: Cmd (the engine also sets CtrlKey with Cmd; MacCtrlKey is the real Control key)
		return 119 === e.KeyCode && e.ShiftKey && !e.AltKey && (isMac ? (e.MacCmdKey && !e.MacCtrlKey) : (e.CtrlKey && !e.MacCmdKey));
	}

	let handleWordKeyDown = CDocument.prototype.handleWordKeyDown;
	CDocument.prototype.handleWordKeyDown = function(e)
	{
		if (isBlockStartKey(e))
		{
			let block = getBlock(this);
			if (block)
				block.byKeys = true;
			else
				this.StartBlockSelection(undefined, undefined, undefined, true);
			return PreventAll;
		}
		let block = getBlock(this);
		if (block)
		{
			let code = e.KeyCode;
			let plain = !e.CtrlKey && !e.AltKey && !e.MacCmdKey;
			if (27 === code)
			{
				collapse(this, block);
				return PreventAll;
			}
			if ((46 === code || 8 === code) && plain && !e.ShiftKey)
			{
				this.DeleteBlockSelection();
				return PreventAll;
			}
			if (block.byKeys && plain && code >= 37 && code <= 40)
			{
				if (37 === code || 39 === code)
					moveFocusX(this, block, 37 === code ? -1 : 1);
				else
					moveFocusY(this, block, 38 === code ? -1 : 1);
				this.ElxUpdateBlock();
				return PreventAll;
			}
			// clipboard and formatting keys keep the block; modifiers alone too
			let isModifier = 16 === code || 17 === code || 18 === code || 91 === code || 92 === code || 93 === code || 224 === code;
			let isCommand  = (e.CtrlKey || e.MacCmdKey) && !e.AltKey;
			if (!isModifier && !isCommand)
				collapse(this, block);
		}
		return handleWordKeyDown ? handleWordKeyDown.call(this, e) : undefined;
	};

	let onKeyPress = CDocument.prototype.OnKeyPress;
	CDocument.prototype.OnKeyPress = function(e)
	{
		let block = getBlock(this);
		if (block)
			collapse(this, block);
		return onKeyPress.apply(this, arguments);
	};

	// ---------------------------------------------------------------- clipboard
	let api = window["Asc"]["asc_docs_api"].prototype;

	let checkCopy = api.asc_CheckCopy;
	api.asc_CheckCopy = function(clipboard, formats)
	{
		let doc = this.private_GetLogicDocument && this.private_GetLogicDocument();
		if (!doc || !getBlock(doc))
			return checkCopy.apply(this, arguments);
		let lines = doc.GetBlockSelectionText() || [];
		let text  = lines.join("\r\n");
		let esc = function(s) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); };
		if (formats & AscCommon.c_oAscClipboardDataFormat.Text)
			clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, text);
		if (formats & AscCommon.c_oAscClipboardDataFormat.Html)
			clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, lines.map(function(l) { return "<p>" + (esc(l) || "&nbsp;") + "</p>"; }).join(""));
	};

	let selectionCut = api.asc_SelectionCut;
	api.asc_SelectionCut = function()
	{
		let doc = this.private_GetLogicDocument && this.private_GetLogicDocument();
		if (doc && getBlock(doc))
		{
			if (!this.isViewMode && !doc.IsViewMode())
				doc.DeleteBlockSelection();
			return;
		}
		return selectionCut.apply(this, arguments);
	};

	api["asc_StartBlockSelection"] = api.asc_StartBlockSelection = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.StartBlockSelection(undefined, undefined, undefined, true) : false;
	};
	api["asc_IsBlockSelection"] = api.asc_IsBlockSelection = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.IsBlockSelection() : false;
	};
	api["asc_GetBlockSelectionText"] = api.asc_GetBlockSelectionText = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.GetBlockSelectionText() : null;
	};
})(window);
