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
 * Word document views and field commands:
 *   Outline view (Ctrl+Alt+O)  paragraphs indented by outline level, outline symbols, no paragraph formatting,
 *                              collapse / expand (Alt+Shift+Minus / Plus), show levels (Alt+Shift+1..9, Alt+Shift+A),
 *                              text formatting on/off (numeric keypad /), Tab / Shift+Tab demote / promote headings
 *   Draft view (Ctrl+Alt+N)    no headers, footers or page margins at the top; floating objects in the text flow
 *   Print Layout (Ctrl+Alt+P)  back to the normal layout
 *   Ctrl+F9                    empty field (or a field whose code is the selected text), its code shown
 *   Ctrl+F11, Ctrl+Shift+F11   lock / unlock fields (a locked field is not updated by F9)
 *   F9                         after the update, a field whose code was shown on its own shows its result again
 * Views are layouts of the engine (word/Editor/Layout); they change neither the document nor its history.
 *   Selection pane           the objects of the document, show / hide (wp:docPr hidden), rename
 *   Reveal Formatting        the formatting at the cursor (asc_GetRevealFormatting)
 * The interface is told through asc_onWordViewChanged(view) and asc_onWordFieldHint(kind, count).
 */
(function(window)
{
	const Type       = Asc.c_oAscDocumentShortcutType;
	const PreventAll = 0xFFFF;

	const OUTLINE_STEP   = 7.5;   // mm of indent per outline level
	const OUTLINE_SYMBOL = 5;     // mm left of the text for the outline symbol
	const BODY_RANK      = 9;     // rank of body text (outline levels are 0..8)
	const SHOW_ALL       = 10;    // show level: everything
	const HIDDEN_LINE    = 1e-6;  // exact line height of a collapsed paragraph (see Paragraph_Recalculate.js)

	const VIEW_PRINT   = "print";
	const VIEW_OUTLINE = "outline";
	const VIEW_DRAFT   = "draft";
	const VIEW_READ    = "read";

	// ---------------------------------------------------------------- layouts
	/**
	 * Outline and Draft layouts: one page geometry for the whole document (the width and the side margins of the
	 * first section), long pages, no headers and footers.
	 * @constructor
	 */
	function CDocumentWordView(logicDocument, kind)
	{
		AscWord.CDocumentReadView.call(this, logicDocument);
		this.Kind = kind;
	}
	CDocumentWordView.prototype = Object.create(AscWord.CDocumentReadView.prototype);
	CDocumentWordView.prototype.constructor = CDocumentWordView;
	CDocumentWordView.prototype.IsReadMode = function()
	{
		return false;
	};
	CDocumentWordView.prototype.GetFontScale = function()
	{
		return 1;
	};
	CDocumentWordView.prototype.IsOutlineView = function()
	{
		return VIEW_OUTLINE === this.Kind;
	};
	CDocumentWordView.prototype.IsWordView = function()
	{
		return true;
	};
	CDocumentWordView.prototype.SetGeometry = function(sectPr)
	{
		let w = sectPr.GetPageWidth(), h = sectPr.GetPageHeight();
		let left = sectPr.GetPageMarginLeft(), right = sectPr.GetPageMarginRight();
		if (VIEW_OUTLINE === this.Kind)
			left = right = Math.min(left, right, 20);

		this.W     = w;
		this.H     = 3 * h;
		this.Scale = 1;

		let viewSectPr = this.SectPr;
		AscCommon.ExecuteNoHistory(function()
		{
			viewSectPr.SetPageSize(w, 3 * h);
			viewSectPr.SetPageMargins(left, 10, right, 10);
		}, this.LogicDocument);
	};

	function getWordView(doc)
	{
		let layout = doc && doc.Layout;
		if (layout && doc.Layouts && layout === doc.Layouts.Read)
			return VIEW_READ;
		return layout && layout.IsWordView && layout.IsWordView() ? layout.Kind : VIEW_PRINT;
	}

	function isOutline(doc)
	{
		return VIEW_OUTLINE === getWordView(doc);
	}

	function recompileAll(doc)
	{
		doc.CheckAllRunContent(function(run)
		{
			run.Recalc_CompiledPr(true);
		});
		doc.GetAllParagraphs().forEach(function(paragraph)
		{
			paragraph.RecalcCompiledPr(true);
		});
		doc.GetAllTables().forEach(function(table)
		{
			table.Recalc_CompiledPr();
		});
	}

	CDocument.prototype.GetWordView = function()
	{
		return getWordView(this);
	};

	/**
	 * @param {string} view "print", "outline", "draft" or "read" (Read Mode: pages of width x height mm, the text
	 *                      flows in them)
	 */
	CDocument.prototype.SetWordView = function(view, width, height)
	{
		if (VIEW_OUTLINE !== view && VIEW_DRAFT !== view && VIEW_READ !== view)
			view = VIEW_PRINT;
		if (view === getWordView(this) && VIEW_READ !== view)
			return;

		if (!this.Layouts.Outline)
		{
			this.Layouts.Outline = new CDocumentWordView(this, VIEW_OUTLINE);
			this.Layouts.Draft   = new CDocumentWordView(this, VIEW_DRAFT);
		}

		let drawingDocument = this.GetDrawingDocument();
		if (drawingDocument)
			drawingDocument.m_bUpdateAllPagesOnFirstRecalculate = true;

		if (VIEW_PRINT === view)
		{
			this.Layout = this.Layouts.Print;
		}
		else if (VIEW_READ === view)
		{
			let read = this.Layouts.Read;
			read.Set(width || 180, height || 240, 1);
			AscCommon.ExecuteNoHistory(function()
			{
				read.SectPr.SetPageMargins(12, 12, 12, 12);
			}, this);
			this.Layout = read;
		}
		else
		{
			let layout = VIEW_OUTLINE === view ? this.Layouts.Outline : this.Layouts.Draft;
			layout.SetGeometry(this.GetSections().GetSectPrByIndex(0));
			this.Layout = layout;
		}

		this.WordOutline = {show : SHOW_ALL, limits : new Map(), hidden : new Set(), plain : false};
		recompileAll(this);
		this.RecalculateFromStart(true);
		this.UpdateSelection();
		this.UpdateInterface();
		this.sendEvent("asc_onWordViewChanged", view);
	};

	// ---------------------------------------------------------------- outline: model
	function headingRank(doc, paraPr)
	{
		let styles = doc.Styles;
		for (let level = 0; level < 9; ++level)
		{
			if (paraPr.PStyle && paraPr.PStyle === styles.GetDefaultHeading(level))
				return level;
		}
		return (undefined !== paraPr.OutlineLvl && null !== paraPr.OutlineLvl && paraPr.OutlineLvl >= 0 && paraPr.OutlineLvl < 9) ? paraPr.OutlineLvl : BODY_RANK;
	}

	function rankOf(paragraph)
	{
		let level = paragraph.GetOutlineLvl();
		return (undefined !== level && null !== level && level >= 0 && level < 9) ? level : BODY_RANK;
	}

	function isTopParagraph(doc, element)
	{
		return element && element.IsParagraph && element.IsParagraph() && element.GetParent() === doc;
	}

	/**
	 * Which paragraphs are hidden by the show level and the collapsed headings
	 */
	CDocument.prototype.private_RefreshOutline = function()
	{
		let state = this.WordOutline;
		let hidden = new Set();
		let stack = [];
		for (let index = 0, count = this.Content.length; index < count; ++index)
		{
			let element = this.Content[index];
			if (!element.IsParagraph())
				continue;

			let rank = rankOf(element);
			while (stack.length && stack[stack.length - 1].rank >= rank)
				stack.pop();

			// text before the first heading follows the show level
			let visible = rank < state.show && (stack.length || undefined === state.orphan || rank < state.orphan);
			for (let i = 0; visible && i < stack.length; ++i)
			{
				if (rank >= stack[i].limit)
					visible = false;
			}
			if (!visible)
				hidden.add(element);

			if (rank < BODY_RANK)
			{
				let limit = state.limits.get(element);
				stack.push({rank : rank, limit : undefined === limit ? SHOW_ALL : limit});
			}
		}
		state.hidden = hidden;

		recompileAll(this);
		this.RecalculateFromStart(true);

		// the cursor leaves the hidden paragraphs for the heading above them
		let current = this.GetCurrentParagraph();
		if (current && hidden.has(current))
		{
			let index = current.GetIndex();
			while (index > 0 && (!this.Content[index].IsParagraph() || hidden.has(this.Content[index])))
				--index;
			this.RemoveSelection();
			this.Content[index].MoveCursorToEndPos(false);
			this.Content[index].Document_SetThisElementCurrent(false);
		}
		this.UpdateSelection();
		this.UpdateInterface();
	};

	CDocument.prototype.IsOutlineHidden = function(paragraph)
	{
		return !!(this.WordOutline && this.WordOutline.hidden.has(paragraph));
	};

	/**
	 * The paragraphs below a heading, up to the next heading of the same or a higher level
	 */
	CDocument.prototype.private_GetOutlineChildren = function(heading)
	{
		let rank = rankOf(heading), result = [];
		for (let index = heading.GetIndex() + 1, count = this.Content.length; index < count; ++index)
		{
			let element = this.Content[index];
			if (!element.IsParagraph())
				continue;
			if (rankOf(element) <= rank)
				break;
			result.push(element);
		}
		return result;
	};

	// headings concerned by expand / collapse: the selected ones, or the one of the current paragraph
	CDocument.prototype.private_GetOutlineTargets = function()
	{
		let paragraphs = this.GetSelectedParagraphs ? this.GetSelectedParagraphs() : [];
		if (!paragraphs || !paragraphs.length)
			paragraphs = [this.GetCurrentParagraph()];

		let result = [];
		for (let i = 0; i < paragraphs.length; ++i)
		{
			let paragraph = paragraphs[i];
			if (!isTopParagraph(this, paragraph))
				continue;
			let index = paragraph.GetIndex();
			while (index >= 0 && (!this.Content[index].IsParagraph() || BODY_RANK === rankOf(this.Content[index])))
				--index;
			if (index >= 0 && -1 === result.indexOf(this.Content[index]))
				result.push(this.Content[index]);
		}
		return result;
	};

	/**
	 * One more (expand) or one less (collapse) level of the text under the heading, like Word
	 */
	CDocument.prototype.OutlineExpandCollapse = function(isExpand)
	{
		let state = this.WordOutline, targets = this.private_GetOutlineTargets();

		// the show level becomes a limit of each heading, so that one heading can be expanded on its own
		if (state.show < SHOW_ALL)
		{
			for (let index = 0, count = this.Content.length; index < count; ++index)
			{
				let element = this.Content[index];
				if (!element.IsParagraph() || rankOf(element) >= state.show)
					continue;
				let limit = state.limits.get(element);
				state.limits.set(element, undefined === limit ? state.show : Math.min(limit, state.show));
			}
			state.orphan = state.show;
			state.show   = SHOW_ALL;
		}

		for (let t = 0; t < targets.length; ++t)
		{
			let heading = targets[t], rank = rankOf(heading);
			let children = this.private_GetOutlineChildren(heading);
			if (!children.length)
				continue;

			let shown = [], hiddenRanks = [];
			for (let i = 0; i < children.length; ++i)
			{
				let childRank = rankOf(children[i]);
				if (state.hidden.has(children[i]))
					hiddenRanks.push(childRank);
				else
					shown.push(childRank);
			}

			if (isExpand)
			{
				if (!hiddenRanks.length)
					continue;
				let next = Math.min.apply(Math, hiddenRanks) + 1;
				if (next > BODY_RANK || !hiddenRanks.some(function(r) { return r >= next; }))
					state.limits.delete(heading);
				else
					state.limits.set(heading, next);
			}
			else
			{
				if (!shown.length)
					continue;
				state.limits.set(heading, Math.max(rank + 1, Math.max.apply(Math, shown)));
			}
		}
		this.private_RefreshOutline();
	};

	/**
	 * Alt+Shift+1..9: headings up to this level; 0 or SHOW_ALL: everything
	 */
	CDocument.prototype.OutlineShowLevel = function(level)
	{
		let state = this.WordOutline;
		state.show = (level >= 1 && level <= 9) ? level : SHOW_ALL;
		state.orphan = undefined;
		state.limits.clear();
		this.private_RefreshOutline();
	};

	// Alt+Shift+A: everything, or the headings only when everything is shown
	CDocument.prototype.OutlineToggleAll = function()
	{
		this.OutlineShowLevel(this.WordOutline.hidden.size ? SHOW_ALL : 9);
	};

	// numeric keypad /: text formatting shown or not
	CDocument.prototype.OutlineToggleFormatting = function()
	{
		this.WordOutline.plain = !this.WordOutline.plain;
		recompileAll(this);
		this.RecalculateFromStart(true);
		this.UpdateSelection();
		this.sendEvent("asc_onWordViewChanged", VIEW_OUTLINE);
	};

	CDocument.prototype.GetOutlineState = function()
	{
		let state = this.WordOutline;
		return {
			"view"           : getWordView(this),
			"showLevel"      : state ? (state.show < SHOW_ALL ? state.show : (undefined !== state.orphan && !state.limits.size ? state.orphan : 0)) : 0,
			"showFormatting" : state ? !state.plain : true,
			"hidden"         : state ? state.hidden.size : 0,
			"level"          : isTopParagraph(this, this.GetCurrentParagraph()) ? rankOf(this.GetCurrentParagraph()) : BODY_RANK
		};
	};

	/**
	 * Promote / demote the current headings (Tab and Shift+Tab of the Outline view, the Outlining tab).
	 * Body text is promoted to the level under the heading above it; demoting Heading 9 makes body text.
	 * @param {number} delta -1 promote, 1 demote, 0 body text, "h1" Heading 1
	 */
	CDocument.prototype.OutlineSetLevel = function(delta)
	{
		let paragraphs = this.GetSelectedParagraphs ? this.GetSelectedParagraphs() : [];
		if (!paragraphs || !paragraphs.length)
			paragraphs = [this.GetCurrentParagraph()];
		paragraphs = paragraphs.filter(function(p) { return p && p.IsParagraph(); });
		if (!paragraphs.length || this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : paragraphs, CheckType : AscCommon.changestype_Paragraph_Properties}))
			return;

		let styles = this.Styles, doc = this;
		this.StartAction(AscDFH.historydescription_Document_SetParagraphStyle);
		paragraphs.forEach(function(paragraph)
		{
			let rank = rankOf(paragraph), next;
			if (delta && "object" === typeof delta)
				next = Math.max(0, Math.min(BODY_RANK, delta.level));
			else if ("h1" === delta)
				next = 0;
			else if (0 === delta)
				next = BODY_RANK;
			else if (BODY_RANK === rank)
			{
				if (delta > 0)
					return;
				let index = paragraph.GetIndex() - 1;
				while (index >= 0 && (!doc.Content[index].IsParagraph() || BODY_RANK === rankOf(doc.Content[index])))
					--index;
				next = index >= 0 && paragraph.GetParent() === doc ? Math.min(rankOf(doc.Content[index]) + 1, 8) : 0;
			}
			else
			{
				next = rank + delta;
				if (next < 0)
					return;
				if (next > 8)
					next = BODY_RANK;
			}
			paragraph.SetParagraphStyleById(BODY_RANK === next ? styles.GetDefaultParagraph() : styles.GetDefaultHeading(next));
		});
		if (isOutline(this))
			recompileAll(this);
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	};

	/**
	 * Alt+Shift+Up / Down in the Outline view: the selected paragraphs (with the text hidden under a collapsed
	 * heading) move above the previous visible paragraph or below the next one (with the text hidden under it)
	 */
	CDocument.prototype.OutlineMoveParagraphs = function(isUp)
	{
		let hidden = this.WordOutline.hidden;
		let paragraph = this.GetCurrentParagraph();
		if (!isTopParagraph(this, paragraph))
			return false;

		let start, end;
		if (this.IsSelectionUse() && docpostype_Content === this.GetDocPosType())
		{
			start = Math.min(this.Selection.StartPos, this.Selection.EndPos);
			end   = Math.max(this.Selection.StartPos, this.Selection.EndPos);
		}
		else
		{
			start = end = paragraph.GetIndex();
		}
		let count = this.Content.length;
		while (end + 1 < count && hidden.has(this.Content[end + 1]))
			++end;

		let from, to;
		if (isUp)
		{
			to = start - 1;
			from = to;
			while (from >= 0 && hidden.has(this.Content[from]))
				--from;
			if (from < 0)
				return true;
		}
		else
		{
			from = end + 1;
			if (from >= count)
				return true;
			to = from;
			while (to + 1 < count && hidden.has(this.Content[to + 1]))
				++to;
		}

		let moved = this.Content.slice(from, to + 1);
		let block = this.Content.slice(start, end + 1);
		if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : moved.concat(block), CheckType : AscCommon.changestype_Paragraph_Content}))
			return true;

		// (Internal_Content_Remove/Add keep CurPos and Selection on the moved paragraphs)
		this.StartAction(AscDFH.historydescription_Document_MoveParagraphByTab);
		this.Internal_Content_Remove(from, moved.length);
		let at = isUp ? from + block.length : start;
		for (let i = 0; i < moved.length; ++i)
			this.Internal_Content_Add(at + i, moved[i]);
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
		return true;
	};

	// ---------------------------------------------------------------- outline: layout and drawing
	let compileParaPr = Paragraph.prototype.Internal_CompileParaPr2;
	Paragraph.prototype.Internal_CompileParaPr2 = function()
	{
		let pr  = compileParaPr.apply(this, arguments);
		let doc = this.LogicDocument;
		if (!doc || !doc.WordOutline || !isOutline(doc) || !isTopParagraph(doc, this) || !pr || !pr.ParaPr)
			return pr;

		let paraPr = pr.ParaPr;
		let rank   = headingRank(doc, paraPr);
		let level  = rank;
		if (BODY_RANK === rank)
		{
			level = 0;
			for (let prev = this.GetPrevParagraph(); prev; prev = prev.GetPrevParagraph())
			{
				if (prev.GetParent() !== doc)
					continue;
				let prevRank = rankOf(prev);
				if (BODY_RANK !== prevRank)
				{
					level = prevRank + 1;
					break;
				}
			}
		}

		// Word does not show the paragraph formatting in the Outline view
		paraPr.Ind.Left      = level * OUTLINE_STEP + OUTLINE_SYMBOL;
		paraPr.Ind.FirstLine = 0;
		paraPr.Ind.Right     = 0;
		paraPr.Jc            = AscCommon.align_Left;
		paraPr.Spacing.Before = 0;
		paraPr.Spacing.After  = BODY_RANK === rank ? 0 : 1.5;
		paraPr.Spacing.BeforeAutoSpacing = false;
		paraPr.Spacing.AfterAutoSpacing  = false;
		paraPr.Spacing.LineRule = Asc.linerule_Auto;
		paraPr.Spacing.Line     = 1;
		paraPr.PageBreakBefore  = false;
		paraPr.ContextualSpacing = false;

		if (doc.WordOutline.hidden.has(this))
		{
			paraPr.Spacing.After    = 0;
			paraPr.Spacing.LineRule = Asc.linerule_Exact;
			paraPr.Spacing.Line     = HIDDEN_LINE;
			paraPr.NumPr            = undefined;
			if (pr.TextPr)
			{
				pr.TextPr.FontSize   = 1;
				pr.TextPr.Vanish     = true;
				pr.TextPr.ElxOutlineHidden = true;
			}
		}
		return pr;
	};

	let compileRunPr = ParaRun.prototype.Internal_Compile_Pr;
	ParaRun.prototype.Internal_Compile_Pr = function()
	{
		let textPr = compileRunPr.apply(this, arguments);
		let paragraph = this.Paragraph;
		let doc = paragraph && paragraph.LogicDocument;
		if (!textPr || !doc || !doc.WordOutline || !isOutline(doc))
			return textPr;

		if (doc.WordOutline.hidden.has(paragraph))
		{
			textPr.FontSize = 1;
			textPr.FontSizeCS = 1;
			textPr.Vanish   = true;
			textPr.ElxOutlineHidden = true;
		}
		else if (doc.WordOutline.plain && isTopParagraph(doc, paragraph))
		{
			// no text formatting: the default size of the document, no bold, italic, underline or colour
			let defaultPr = doc.Styles.Default.TextPr;
			textPr.Bold      = false;
			textPr.Italic    = false;
			textPr.Underline = false;
			textPr.Strikeout = false;
			textPr.DStrikeout = false;
			textPr.Caps      = false;
			textPr.SmallCaps = false;
			textPr.Color     = new CDocumentColor(0, 0, 0, true);
			textPr.Unifill   = undefined;
			textPr.HighLight = highlight_None;
			if (defaultPr && defaultPr.FontSize)
				textPr.FontSize = textPr.FontSizeCS = defaultPr.FontSize;
		}
		return textPr;
	};

	let isHiddenText = AscWord.IsHiddenText;
	AscWord.IsHiddenText = function(textPr)
	{
		return !!(textPr && textPr.ElxOutlineHidden) || isHiddenText(textPr);
	};

	let drawParagraph = Paragraph.prototype.Draw;
	Paragraph.prototype.Draw = function(curPage, graphics)
	{
		let doc = this.LogicDocument;
		if (doc && doc.WordOutline && doc.WordOutline.hidden.size && doc.WordOutline.hidden.has(this) && isOutline(doc))
			return;
		drawParagraph.call(this, curPage, graphics);
	};

	// outline symbols left of the text: plus for a heading with text under it, minus for a heading without, a small
	// circle for body text
	let drawParagraph2 = Paragraph.prototype.Internal_Draw_2;
	Paragraph.prototype.Internal_Draw_2 = function(curPage, graphics, pr)
	{
		drawParagraph2.call(this, curPage, graphics, pr);

		let doc = this.LogicDocument;
		if (!doc || !isOutline(doc) || !isTopParagraph(doc, this) || !this.Check_FirstPage(curPage) || !graphics.isSupportEditFeatures || !graphics.isSupportEditFeatures())
			return;

		let page = this.Pages[curPage];
		let line = this.Lines[page.FirstLine];
		if (!line || !line.Ranges[0])
			return;

		let rank     = rankOf(this);
		let baseline = page.Y + line.Y;
		let height   = Math.max(1.5, Math.min(4, (line.Metrics.TextAscent || 3) * 0.75));
		let cx       = line.Ranges[0].X - OUTLINE_SYMBOL / 2 - 0.5;
		let cy       = baseline - height / 2;
		let next     = this.GetNextParagraph();
		while (next && next.GetParent() !== doc)
			next = next.GetNextParagraph();

		let penW = 0.25;
		graphics.p_color(96, 96, 96, 255);
		graphics.b_color1(96, 96, 96, 255);
		graphics.p_width(penW * 1000);
		if (BODY_RANK === rank)
		{
			drawCircle(graphics, cx, cy, 0.7);
			graphics.ds();
			return;
		}

		let r = 1.6;
		drawCircle(graphics, cx, cy, r);
		graphics.ds();
		graphics._s();
		graphics._m(cx - r + 0.6, cy);
		graphics._l(cx + r - 0.6, cy);
		graphics.ds();
		if (next && rankOf(next) > rank)
		{
			graphics._s();
			graphics._m(cx, cy - r + 0.6);
			graphics._l(cx, cy + r - 0.6);
			graphics.ds();
		}
		// a collapsed heading is underlined in grey
		if (next && doc.WordOutline && doc.WordOutline.hidden.has(next))
		{
			let x0 = line.Ranges[0].X, x1 = line.Ranges[0].X + line.Ranges[0].W;
			graphics.p_color(160, 160, 160, 255);
			graphics._s();
			graphics._m(x0, baseline + 1);
			graphics._l(Math.max(x1, x0 + 10), baseline + 1);
			graphics.ds();
		}
	};

	function drawCircle(graphics, cx, cy, r)
	{
		let k = 0.5523 * r;
		graphics._s();
		graphics._m(cx + r, cy);
		graphics._c(cx + r, cy + k, cx + k, cy + r, cx, cy + r);
		graphics._c(cx - k, cy + r, cx - r, cy + k, cx - r, cy);
		graphics._c(cx - r, cy - k, cx - k, cy - r, cx, cy - r);
		graphics._c(cx + k, cy - r, cx + r, cy - k, cx + r, cy);
		graphics._z();
	}

	// the Outline and Draft views show floating objects in the text flow
	let isInline = ParaDrawing.prototype.Is_Inline;
	ParaDrawing.prototype.Is_Inline = function()
	{
		let paragraph = this.GetParagraph ? this.GetParagraph() : null;
		let doc = paragraph && paragraph.LogicDocument;
		if (doc && doc.Layout && doc.Layout.IsWordView && doc.Layout.IsWordView() && paragraph.bFromDocument && !(paragraph.Parent && paragraph.Parent.IsHdrFtr && paragraph.Parent.IsHdrFtr()))
			return true;
		return isInline.apply(this, arguments);
	};

	// ---------------------------------------------------------------- fields
	/**
	 * Ctrl+F9: an empty field with its code shown, the cursor inside; a selection becomes the field code
	 */
	CDocument.prototype.InsertEmptyField = function()
	{
		if (docpostype_DrawingObjects === this.GetDocPosType() && !this.GetCurrentParagraph())
			return;
		if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			return;

		let code = "";
		if (this.IsSelectionUse() && !this.IsSelectionEmpty(true))
			code = (this.GetSelectedText(true) || "").replace(/[\r\n]+/g, " ");

		this.StartAction(AscDFH.historydescription_Document_AddComplexField);
		if (code)
			this.Remove(1, false, false, true);
		let textPr = this.GetDirectTextPr().Copy();
		// no result yet: Word computes it with F9
		this.ElxNoFieldUpdate = true;
		let field = this.AddFieldWithInstruction(" " + code + " ", textPr);
		this.ElxNoFieldUpdate = false;
		if (field && field.GetBeginChar())
			field.GetBeginChar().SetShowFieldCode(true);
		this.Recalculate();
		if (field && field.GetBeginChar())
		{
			field.GetBeginChar().MoveCursorToChar(false);
			this.MoveCursorRight(false, false);
			if (code)
				for (let i = 0; i < code.length; ++i)
					this.MoveCursorRight(false, false);
		}
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.private_GetFieldsForCommand = function()
	{
		let fields = this.IsSelectionUse() && !this.IsSelectionEmpty(true) ? this.GetAllFields(true) : this.GetCurrentComplexFields();
		return (fields || []).filter(function(field) { return field instanceof CComplexField && field.IsValid() && field.GetBeginChar(); });
	};

	/**
	 * Ctrl+F11 / Ctrl+Shift+F11: lock or unlock the selected fields, or the field of the cursor
	 */
	CDocument.prototype.SetFieldsLocked = function(isLocked)
	{
		let all = this.private_GetFieldsForCommand();
		let fields = all.filter(function(field)
		{
			return !!field.IsLocked() !== !!isLocked && field.GetBeginChar().GetRun();
		});
		let changed = fields.length;
		if (changed)
		{
			let runs = fields.map(function(field) { return field.GetBeginChar().GetRun(); });
			if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : runs.map(function(run) { return run.GetParagraph(); }), CheckType : AscCommon.changestype_Paragraph_Content}))
				return 0;
			// The lock is a property of the begin character (w:fldLock): the character is replaced through the
			// history, so that the lock is undone, sent to co-editors and saved.
			this.StartAction(AscDFH.historydescription_Document_AddComplexField);
			let doc = this;
			fields.forEach(function(field)
			{
				let oldChar = field.GetBeginChar();
				let run = oldChar.GetRun();
				let pos = run.GetElementPosition(oldChar);
				if (pos < 0)
					return;
				// both characters belong to the field: the one in the run is its begin character (see SetParent below,
				// for undo and redo)
				let newChar = new ParaFieldChar(fldchartype_Begin, doc);
				newChar.ComplexField = field;
				newChar.ElxLockSwap = oldChar.ElxLockSwap = true;
				newChar.showFieldCode = oldChar.showFieldCode;
				newChar.fldData = oldChar.fldData;
				newChar.ffData = oldChar.ffData ? oldChar.ffData.Copy() : null;
				newChar.ElxFieldLock = !!isLocked;
				run.RemoveFromContent(pos, 1, true);
				run.AddToContent(pos, newChar, true);
				field.ReplaceChar(newChar);
			});
			this.FinalizeAction();
		}
		this.sendEvent("asc_onWordFieldHint", all.length ? (isLocked ? "locked" : "unlocked") : "nofield", all.length);
		return all.length;
	};

	let setFieldCharParent = ParaFieldChar.prototype.SetParent;
	ParaFieldChar.prototype.SetParent = function(parent)
	{
		setFieldCharParent.apply(this, arguments);
		if (this.ElxLockSwap && parent && this.IsBegin() && this.ComplexField)
			this.ComplexField.ReplaceChar(this);
	};

	CComplexField.prototype.IsLocked = function()
	{
		return !!(this.BeginChar && this.BeginChar.ElxFieldLock);
	};

	let updateComplexField = CComplexField.prototype.Update;
	CComplexField.prototype.Update = function()
	{
		if (this.IsLocked() || (this.LogicDocument && this.LogicDocument.ElxNoFieldUpdate))
			return false;
		return updateComplexField.apply(this, arguments);
	};

	// A shown field code is drawn through replacement text items; the instruction is still made of the
	// instruction characters, so that a code typed or edited in place is used by the next update.
	ParaRun.prototype.ProcessComplexFields = function(complexFields)
	{
		let isRemovedInReview = (reviewtype_Remove === this.GetReviewType());
		for (let pos = 0, count = this.Content.length; pos < count; ++pos)
		{
			let item = this.Content[pos];
			let isInstr = para_InstrText === item.Type;
			if (!isInstr)
				item = this.private_CheckInstrText(item);

			if (complexFields.isHiddenFieldContent() && para_End !== item.Type && para_FieldChar !== item.Type)
				continue;

			if (para_FieldChar === item.Type)
				complexFields.processFieldCharAndCollectComplexField(item);
			else if (isInstr && !isRemovedInReview)
				complexFields.processInstruction(item);
		}
	};

	let updateFields = CDocument.prototype.UpdateFields;
	CDocument.prototype.UpdateFields = function(isBySelection)
	{
		let fields = (this.GetAllFields(isBySelection) || []).concat(this.GetCurrentComplexFields() || []).filter(function(field)
		{
			return field instanceof CComplexField && field.IsValid();
		});
		// a code typed or edited in place (Ctrl+F9, Shift+F9) is read again before the update; only for a shown
		// code whose field lies in one paragraph (reading one paragraph of a longer field, a TOC, would cut it)
		let paragraphs = [];
		fields.forEach(function(field)
		{
			if (!field.IsShowFieldCode() || !field.GetEndChar())
				return;
			let run = field.GetBeginChar() && field.GetBeginChar().GetRun();
			let endRun = field.GetEndChar().GetRun();
			let paragraph = run && run.GetParagraph();
			if (paragraph && endRun && endRun.GetParagraph() === paragraph && -1 === paragraphs.indexOf(paragraph))
				paragraphs.push(paragraph);
		});
		paragraphs.forEach(function(paragraph) { paragraph.ProcessComplexFields(); });

		let locked = fields.filter(function(field) { return field.IsLocked(); });
		let all    = this.GetAllFields(false) || [];
		let allShown = all.length > 1 && all.every(function(field) { return !(field instanceof CComplexField) || field.IsShowFieldCode(); });

		updateFields.apply(this, arguments);

		// a field whose code was shown on its own (Ctrl+F9, Shift+F9) shows its result after the update, like Word
		let shown = allShown ? [] : fields.filter(function(field) { return !field.IsLocked() && field.IsShowFieldCode(); });
		if (shown.length)
		{
			shown.forEach(function(field) { field.GetBeginChar().SetShowFieldCode(false); });
			this.Recalculate();
			let last = shown[shown.length - 1];
			if (last.GetEndChar())
				last.GetEndChar().MoveCursorToChar(false);
			this.UpdateSelection();
		}
		if (locked.length)
			this.sendEvent("asc_onWordFieldHint", "lockedskip", locked.length);
	};

	// ---------------------------------------------------------------- shortcuts
	let executeWordShortcut = CDocument.prototype.executeWordShortcut;
	CDocument.prototype.executeWordShortcut = function(type)
	{
		switch (type)
		{
			case Type.InsertEmptyField:
				this.InsertEmptyField();
				return true;
			case Type.LockField:
				this.SetFieldsLocked(true);
				return true;
			case Type.UnlockField:
				this.SetFieldsLocked(false);
				return true;
			case Type.OutlineView:
				this.SetWordView(VIEW_OUTLINE);
				return true;
			case Type.DraftView:
				this.SetWordView(VIEW_DRAFT);
				return true;
			case Type.PrintLayoutView:
				this.SetWordView(VIEW_PRINT);
				this.sendEvent("asc_onWordViewChanged", VIEW_PRINT);
				return true;
			case Type.MoveParagraphUp:
			case Type.MoveParagraphDown:
				if (isOutline(this) && this.WordOutline && this.OutlineMoveParagraphs(Type.MoveParagraphUp === type))
					return true;
				break;
			case Type.PromoteHeading:
			case Type.DemoteHeading:
				if (isOutline(this))
				{
					this.OutlineSetLevel(Type.PromoteHeading === type ? -1 : 1);
					return true;
				}
				break;
		}
		return executeWordShortcut ? executeWordShortcut.call(this, type) : false;
	};

	function isPlus(e)
	{
		return 187 === e.KeyCode || 61 === e.KeyCode || 107 === e.KeyCode;
	}

	function isMinus(e)
	{
		return 189 === e.KeyCode || 173 === e.KeyCode || 109 === e.KeyCode;
	}

	// keys of the Outline view (Windows: Alt+Shift+..., macOS: Control+Shift+...)
	let handleWordKeyDown = CDocument.prototype.handleWordKeyDown;
	CDocument.prototype.handleWordKeyDown = function(e)
	{
		if (this.WordOutline && isOutline(this) && docpostype_Content === this.GetDocPosType())
		{
			let code  = e.KeyCode;
			let isMac = AscCommon.AscBrowser.isMacOs;
			let outlineMod = e.ShiftKey && !e.MacCmdKey && (isMac ? (e.CtrlKey && !e.AltKey) : (e.AltKey && !e.CtrlKey));
			if (outlineMod)
			{
				if (isPlus(e) || isMinus(e))
				{
					this.OutlineExpandCollapse(isPlus(e));
					return PreventAll;
				}
				if (code >= 49 && code <= 57)
				{
					this.OutlineShowLevel(code - 48);
					return PreventAll;
				}
				if (65 === code)
				{
					this.OutlineToggleAll();
					return PreventAll;
				}
			}
			let noMod = !e.CtrlKey && !e.AltKey && !e.MacCmdKey;
			if (111 === code && noMod && !e.ShiftKey)
			{
				this.OutlineToggleFormatting();
				return PreventAll;
			}
			if (9 === code && noMod)
			{
				let paragraph = this.GetCurrentParagraph();
				if (isTopParagraph(this, paragraph) && (BODY_RANK !== rankOf(paragraph) || e.ShiftKey))
				{
					this.OutlineSetLevel(e.ShiftKey ? -1 : 1);
					return PreventAll;
				}
			}
			// the cursor does not stop in collapsed text
			if (this.WordOutline.hidden.size && !e.CtrlKey && !e.AltKey && !e.MacCmdKey && (37 === code || 38 === code || 39 === code || 40 === code))
			{
				let forward = 39 === code || 40 === code;
				let move = function(doc)
				{
					if (37 === code)
						doc.MoveCursorLeft(e.ShiftKey, false);
					else if (39 === code)
						doc.MoveCursorRight(e.ShiftKey, false);
					else if (38 === code)
						doc.MoveCursorUp(e.ShiftKey, false);
					else
						doc.MoveCursorDown(e.ShiftKey, false);
				};
				move(this);
				for (let guard = 0; guard < 100000; ++guard)
				{
					let paragraph = this.GetCurrentParagraph();
					if (!paragraph || !this.WordOutline.hidden.has(paragraph))
						break;
					let index = paragraph.GetIndex();
					// jump over the whole hidden block
					let target = index;
					while (this.Content[target] && this.WordOutline.hidden.has(this.Content[target]))
						target += forward ? 1 : -1;
					if (!this.Content[target])
					{
						target -= forward ? 1 : -1;
						while (target >= 0 && target < this.Content.length && this.WordOutline.hidden.has(this.Content[target]))
							target += forward ? -1 : 1;
					}
					if (target < 0 || target >= this.Content.length)
						break;
					let element = this.Content[target];
					if (e.ShiftKey)
					{
						move(this);
						continue;
					}
					this.RemoveSelection();
					if (forward)
						element.MoveCursorToStartPos(false);
					else
						element.MoveCursorToEndPos(false);
					element.Document_SetThisElementCurrent(false);
					break;
				}
				this.UpdateSelection();
				this.UpdateInterface();
				return PreventAll;
			}
		}
		return handleWordKeyDown ? handleWordKeyDown.call(this, e) : undefined;
	};

	// ---------------------------------------------------------------- default keys
	(function()
	{
		const T = Type;
		const keyCodes = Asc.c_oAscKeyCodes;
		const isMac = AscCommon.AscBrowser.isMacOs;
		// [type, Windows/Linux keys, macOS keys]; a key is [keyCode, ctrl, shift, alt, cmd]
		const bindings = [
			[T.InsertEmptyField, [[keyCodes.F9, 1]],          [[keyCodes.F9, 0, 0, 0, 1]]],
			[T.LockField,        [[keyCodes.F11, 1]],         [[keyCodes.F11, 0, 0, 0, 1]]],
			[T.UnlockField,      [[keyCodes.F11, 1, 1]],      [[keyCodes.F11, 0, 1, 0, 1]]],
			[T.OutlineView,      [[keyCodes.KeyO, 1, 0, 1]],  [[keyCodes.KeyO, 0, 0, 1, 1]]],
			[T.DraftView,        [[keyCodes.KeyN, 1, 0, 1]],  [[keyCodes.KeyN, 0, 0, 1, 1]]],
			[T.PrintLayoutView,  [[keyCodes.KeyP, 1, 0, 1]],  [[keyCodes.KeyP, 0, 0, 1, 1]]]
		];
		for (let i = 0; i < bindings.length; ++i)
		{
			let type = bindings[i][0];
			let keys = bindings[i][isMac ? 2 : 1];
			Asc.c_oAscDefaultShortcuts[type] = keys.map(function(k) { return new Asc.CAscShortcut(type, k[0], !!k[1], !!k[2], !!k[3], !!k[4]); });
			Asc.c_oAscUnlockedShortcutActionTypes[type] = true;
		}
	})();

	// ---------------------------------------------------------------- api
	let api = window["Asc"]["asc_docs_api"].prototype;
	api["asc_SetWordView"] = api.asc_SetWordView = function(view, width, height)
	{
		let doc = this.private_GetLogicDocument();
		doc && doc.SetWordView(view, width, height);
	};
	api["asc_GetWordView"] = api.asc_GetWordView = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.GetWordView() : VIEW_PRINT;
	};
	api["asc_GetOutlineState"] = api.asc_GetOutlineState = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.GetOutlineState() : null;
	};
	/**
	 * Commands of the Outlining tab: "promote", "demote", "body", "h1", "setlevel" (0..8, 9 body text), "up", "down", "expand", "collapse",
	 * "level" (with the level, 0 for all), "formatting"
	 */
	api["asc_OutlineCommand"] = api.asc_OutlineCommand = function(command, value)
	{
		let doc = this.private_GetLogicDocument();
		if (!doc || !isOutline(doc))
			return;
		switch (command)
		{
			case "promote": doc.OutlineSetLevel(-1); break;
			case "demote": doc.OutlineSetLevel(1); break;
			case "body": doc.OutlineSetLevel(0); break;
			case "h1": doc.OutlineSetLevel("h1"); break;
			case "setlevel": doc.OutlineSetLevel({level : value}); break;
			case "up": doc.executeWordShortcut(Type.MoveParagraphUp); break;
			case "down": doc.executeWordShortcut(Type.MoveParagraphDown); break;
			case "expand": doc.OutlineExpandCollapse(true); break;
			case "collapse": doc.OutlineExpandCollapse(false); break;
			case "level": doc.OutlineShowLevel(value); break;
			case "formatting": doc.OutlineToggleFormatting(); break;
		}
		this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.UpdateSelection();
	};
	api["asc_InsertEmptyField"] = api.asc_InsertEmptyField = function()
	{
		let doc = this.private_GetLogicDocument();
		doc && doc.InsertEmptyField();
	};
	api["asc_SetFieldsLocked"] = api.asc_SetFieldsLocked = function(isLocked)
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.SetFieldsLocked(isLocked) : 0;
	};

	/**
	 * Reveal Formatting (Shift+F1): the formatting at the cursor or of the selection, in millimetres and points.
	 * Values that differ in the selection are null.
	 */
	CDocument.prototype.GetRevealFormatting = function()
	{
		let textPr = this.GetCalculatedTextPr() || new CTextPr();
		let paraPr = this.GetCalculatedParaPr() || new CParaPr();
		let sectPr = this.GetCurrentSectPr ? this.GetCurrentSectPr() : null;
		let style  = paraPr.PStyle ? this.Styles.Get(paraPr.PStyle) : this.Styles.Get(this.Styles.GetDefaultParagraph());

		let color = null;
		try
		{
			if (textPr.Unifill && textPr.Unifill.fill)
			{
				textPr.Unifill.check(this.GetTheme(), this.GetColorMap());
				let rgba = textPr.Unifill.getRGBAColor();
				color = {r : rgba.R, g : rgba.G, b : rgba.B, auto : false};
			}
			else if (textPr.Color)
				color = {r : textPr.Color.r, g : textPr.Color.g, b : textPr.Color.b, auto : !!textPr.Color.Auto};
		}
		catch (e)
		{
			color = null;
		}

		let font = textPr.RFonts && textPr.RFonts.Ascii ? textPr.RFonts.Ascii.Name : (textPr.FontFamily ? textPr.FontFamily.Name : null);
		let numPr = paraPr.NumPr;
		let result = {
			"font" : {
				"name"      : font || null,
				"size"      : undefined !== textPr.FontSize ? textPr.FontSize : null,
				"bold"      : undefined !== textPr.Bold ? !!textPr.Bold : null,
				"italic"    : undefined !== textPr.Italic ? !!textPr.Italic : null,
				"underline" : undefined !== textPr.Underline ? !!textPr.Underline : null,
				"strikeout" : undefined !== textPr.Strikeout ? !!(textPr.Strikeout || textPr.DStrikeout) : null,
				"caps"      : undefined !== textPr.Caps ? !!textPr.Caps : null,
				"smallCaps" : undefined !== textPr.SmallCaps ? !!textPr.SmallCaps : null,
				"hidden"    : !!textPr.Vanish,
				"position"  : undefined !== textPr.VertAlign ? textPr.VertAlign : null,
				"spacing"   : undefined !== textPr.Spacing ? textPr.Spacing : null,
				"highlight" : textPr.HighLight && highlight_None !== textPr.HighLight ? {r : textPr.HighLight.r, g : textPr.HighLight.g, b : textPr.HighLight.b} : null,
				"color"     : color,
				"lang"      : textPr.Lang && undefined !== textPr.Lang.Val ? textPr.Lang.Val : null
			},
			"paragraph" : {
				"style"           : style ? style.GetName() : null,
				"align"           : undefined !== paraPr.Jc ? paraPr.Jc : null,
				"outlineLevel"    : undefined !== paraPr.OutlineLvl && null !== paraPr.OutlineLvl ? paraPr.OutlineLvl : null,
				"left"            : paraPr.Ind && undefined !== paraPr.Ind.Left ? paraPr.Ind.Left : null,
				"right"           : paraPr.Ind && undefined !== paraPr.Ind.Right ? paraPr.Ind.Right : null,
				"firstLine"       : paraPr.Ind && undefined !== paraPr.Ind.FirstLine ? paraPr.Ind.FirstLine : null,
				"before"          : paraPr.Spacing && undefined !== paraPr.Spacing.Before ? paraPr.Spacing.Before : null,
				"after"           : paraPr.Spacing && undefined !== paraPr.Spacing.After ? paraPr.Spacing.After : null,
				"line"            : paraPr.Spacing && undefined !== paraPr.Spacing.Line ? paraPr.Spacing.Line : null,
				"lineRule"        : paraPr.Spacing && undefined !== paraPr.Spacing.LineRule ? paraPr.Spacing.LineRule : null,
				"keepNext"        : !!paraPr.KeepNext,
				"keepLines"       : !!paraPr.KeepLines,
				"pageBreakBefore" : !!paraPr.PageBreakBefore,
				"widowControl"    : !!paraPr.WidowControl,
				"numbered"        : !!(numPr && numPr.NumId && "0" !== numPr.NumId)
			},
			"section" : null
		};
		if (sectPr)
		{
			result["section"] = {
				"width"   : sectPr.GetPageWidth(),
				"height"  : sectPr.GetPageHeight(),
				"left"    : sectPr.GetPageMarginLeft(),
				"right"   : sectPr.GetPageMarginRight(),
				"top"     : sectPr.GetPageMarginTop(),
				"bottom"  : sectPr.GetPageMarginBottom(),
				"columns" : sectPr.GetColumnCount ? sectPr.GetColumnCount() : 1
			};
		}
		return result;
	};
	api["asc_GetRevealFormatting"] = api.asc_GetRevealFormatting = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.GetRevealFormatting() : null;
	};

	// ---------------------------------------------------------------- Selection pane (Alt+F10)
	// A hidden object (wp:docPr hidden) is kept in the document, not drawn, and the text does not wrap around it.
	function isHiddenDrawing(paraDrawing)
	{
		return !!(paraDrawing && paraDrawing.docPr && paraDrawing.docPr.isHidden);
	}

	let useTextWrap = ParaDrawing.prototype.Use_TextWrap;
	ParaDrawing.prototype.Use_TextWrap = function()
	{
		return isHiddenDrawing(this) ? false : useTextWrap.apply(this, arguments);
	};

	["CShape", "CImageShape", "CGroupShape", "CChartSpace", "CGraphicFrame", "SmartArt"].forEach(function(name)
	{
		let cls = AscFormat[name];
		if (!cls || !cls.prototype || !cls.prototype.draw)
			return;
		let draw = cls.prototype.draw;
		cls.prototype.draw = function()
		{
			if (this.parent instanceof ParaDrawing && isHiddenDrawing(this.parent))
				return;
			return draw.apply(this, arguments);
		};
	});

	const DRAWING_KIND = {};
	DRAWING_KIND[AscDFH.historyitem_type_ImageShape] = "picture";
	DRAWING_KIND[AscDFH.historyitem_type_Shape]      = "shape";
	DRAWING_KIND[AscDFH.historyitem_type_GroupShape] = "group";
	DRAWING_KIND[AscDFH.historyitem_type_ChartSpace] = "chart";
	DRAWING_KIND[AscDFH.historyitem_type_SmartArt]   = "smartart";

	CDocument.prototype.private_GetPaneDrawings = function()
	{
		return this.GetAllDrawingObjects().filter(function(drawing)
		{
			return drawing.GraphicObj && drawing.IsUseInDocument && drawing.IsUseInDocument();
		});
	};

	/**
	 * The objects of the document for the Selection pane, the one in front first (as Word lists them)
	 */
	CDocument.prototype.GetSelectionPaneObjects = function()
	{
		let selected = this.DrawingObjects ? this.DrawingObjects.getSelectedObjects() : [];
		let counters = {};
		let list = this.private_GetPaneDrawings().map(function(drawing)
		{
			let graphic = drawing.GraphicObj;
			let kind = DRAWING_KIND[graphic.getObjectType()] || "object";
			if ("shape" === kind && graphic.txBody === undefined && graphic.textBoxContent)
				kind = "textbox";
			counters[kind] = (counters[kind] || 0) + 1;
			return {
				"id"       : drawing.GetId(),
				"name"     : drawing.docPr && drawing.docPr.name ? drawing.docPr.name : null,
				"kind"     : kind,
				"number"   : counters[kind],
				"hidden"   : isHiddenDrawing(drawing),
				"inline"   : drawing.IsInline(),
				"page"     : drawing.GetPageNum ? drawing.GetPageNum() : 0,
				"z"        : drawing.IsInline() ? -1 : (drawing.RelativeHeight || 0),
				"selected" : selected.indexOf(graphic) !== -1
			};
		});
		list.sort(function(a, b) { return a["page"] - b["page"] || b["z"] - a["z"]; });
		return list;
	};

	CDocument.prototype.private_GetPaneDrawing = function(id)
	{
		let drawing = AscCommon.g_oTableId.Get_ById(id);
		return drawing instanceof ParaDrawing ? drawing : null;
	};

	CDocument.prototype.SelectPaneObject = function(id)
	{
		let drawing = this.private_GetPaneDrawing(id);
		let paragraph = drawing && drawing.GetParagraph();
		if (!paragraph || isHiddenDrawing(drawing))
			return false;
		this.SelectDrawings([drawing], paragraph.GetParent());
		this.UpdateSelection();
		this.UpdateInterface();
		return true;
	};

	/**
	 * Show or hide objects (one undo step); ids: one id or a list, or null for all objects
	 */
	CDocument.prototype.SetPaneObjectsHidden = function(ids, isHidden)
	{
		let doc = this;
		let drawings = null === ids ? this.private_GetPaneDrawings() : [].concat(ids).map(function(id) { return doc.private_GetPaneDrawing(id); }).filter(Boolean);
		drawings = drawings.filter(function(drawing) { return isHiddenDrawing(drawing) !== isHidden; });
		if (!drawings.length)
			return;
		let paragraphs = drawings.map(function(drawing) { return drawing.GetParagraph(); }).filter(Boolean);
		if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : paragraphs, CheckType : AscCommon.changestype_Paragraph_Content}))
			return;

		this.StartAction(AscDFH.historydescription_Document_ApplyImagePr);
		if (isHidden && this.DrawingObjects)
		{
			let selected = this.DrawingObjects.getSelectedObjects();
			if (drawings.some(function(drawing) { return selected.indexOf(drawing.GraphicObj) !== -1; }))
			{
				this.RemoveSelection();
				this.SetDocPosType(docpostype_Content);
			}
		}
		drawings.forEach(function(drawing) { drawing.docPr.setIsHidden(isHidden ? true : null); });
		this.Recalculate();
		this.GetDrawingDocument().ClearCachePages && this.GetDrawingDocument().ClearCachePages();
		this.GetDrawingDocument().FirePaint && this.GetDrawingDocument().FirePaint();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	};

	CDocument.prototype.RenamePaneObject = function(id, name)
	{
		let drawing = this.private_GetPaneDrawing(id);
		if (!drawing || !name || !drawing.GetParagraph())
			return;
		if (this.IsSelectionLocked(AscCommon.changestype_None, {Type : AscCommon.changestype_2_Element_and_Type, Element : drawing.GetParagraph(), CheckType : AscCommon.changestype_Paragraph_Content}))
			return;
		this.StartAction(AscDFH.historydescription_Document_ApplyImagePr);
		drawing.docPr.setName(name);
		this.FinalizeAction();
	};

	api["asc_GetSelectionPaneObjects"] = api.asc_GetSelectionPaneObjects = function()
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.GetSelectionPaneObjects() : [];
	};
	api["asc_SelectPaneObject"] = api.asc_SelectPaneObject = function(id)
	{
		let doc = this.private_GetLogicDocument();
		return doc ? doc.SelectPaneObject(id) : false;
	};
	api["asc_SetPaneObjectsHidden"] = api.asc_SetPaneObjectsHidden = function(ids, isHidden)
	{
		let doc = this.private_GetLogicDocument();
		doc && doc.SetPaneObjectsHidden(ids, isHidden);
	};
	api["asc_RenamePaneObject"] = api.asc_RenamePaneObject = function(id, name)
	{
		let doc = this.private_GetLogicDocument();
		doc && doc.RenamePaneObject(id, name);
	};

	AscWord.CDocumentWordView = CDocumentWordView;
})(window);
