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

var tab_Bar     = Asc.c_oAscTabType.Bar;
var tab_Clear   = Asc.c_oAscTabType.Clear;
var tab_Decimal = Asc.c_oAscTabType.Decimal;
var tab_Num     = Asc.c_oAscTabType.Num;
var tab_Right   = Asc.c_oAscTabType.Right;
var tab_Left    = Asc.c_oAscTabType.Left;
var tab_Center  = Asc.c_oAscTabType.Center;

let tab_Symbol     = 0x0022;//0x2192;
let tab_Symbol_Rtl = 0x0021;

(function(window)
{
	let tabWidth = null;
	
	function getTabActualWidth(tabCode)
	{
		if (null !== tabWidth)
			return tabWidth;
		
		g_oTextMeasurer.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});
		tabWidth = g_oTextMeasurer.Measure(String.fromCharCode(tabCode)).Width;
		return tabWidth;
	}

	// TODO: Реализовать табы по точке и с чертой (tab_Bar tab_Decimal)


	/**
	 * Класс представляющий элемент табуляции.
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunTab()
	{
		AscWord.CRunElementBase.call(this);
		
		this.Width       = 0;
		this.LeaderCode  = 0x00;
		this.LeaderWidth = 0;
	}
	CRunTab.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunTab.prototype.constructor = CRunTab;

	CRunTab.prototype.Type = para_Tab;
	CRunTab.prototype.IsTab = function()
	{
		return true;
	};
	CRunTab.prototype.Draw = function(X, Y, Context, drawState)
	{
		if (this.Width > 0.01 && 0 !== this.LeaderCode)
		{
			Context.SetFontSlot(AscWord.fontslot_ASCII, 1);
			var nCharsCount = Math.floor(this.Width / this.LeaderWidth);
			
			var _X = X + (this.Width - nCharsCount * this.LeaderWidth) / 2;
			for (var nIndex = 0; nIndex < nCharsCount; ++nIndex, _X += this.LeaderWidth)
				Context.FillTextCode(_X, Y, this.LeaderCode);
		}

		if(Context.m_bIsTextDrawer)
		{
			Context.CheckSpaceDraw(this);
		}
		
		if (editor && editor.ShowParaMarks)
		{
			let isRtl = drawState.isRtlMainDirection();
			
			Context.p_color(0, 0, 0, 255);
			Context.b_color1(0, 0, 0, 255);
			
			let tabCode = drawState.isRtlMainDirection() ? tab_Symbol_Rtl : tab_Symbol;
			let tabActualWidth = getTabActualWidth(tabCode);
			var X0 = this.Width / 2 - tabActualWidth / 2;

			Context.SetFont({FontFamily : {Name : "ASCW3", Index : -1}, FontSize : 10, Italic : false, Bold : false});

			if (X0 > 0)
				Context.FillText2(X + X0, Y, String.fromCharCode(tabCode), 0, this.Width);
			else
				Context.FillText2(X, Y, String.fromCharCode(tabCode), isRtl ? 0 : tabActualWidth - this.Width, this.Width);
		}
	};
	CRunTab.prototype.Measure = function(Context)
	{
	};
	CRunTab.prototype.SetLeader = function(leaderType, textPr)
	{
		g_oTextMeasurer.SetTextPr(textPr);
		g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, 1);
		
		switch (leaderType)
		{
			case Asc.c_oAscTabLeader.Dot: // "."
				this.LeaderCode  = 0x002E;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x002E).Width;
				break;
			case Asc.c_oAscTabLeader.Heavy: // "_"
			case Asc.c_oAscTabLeader.Underscore:
				this.LeaderCode  = 0x005F;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x005F).Width;
				break;
			case Asc.c_oAscTabLeader.Hyphen: // "-"
				this.LeaderCode  = 0x002D;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x002D).Width * 1.5;
				break;
			case Asc.c_oAscTabLeader.MiddleDot: // "·"
				this.LeaderCode  = 0x00B7;
				this.LeaderWidth = g_oTextMeasurer.MeasureCode(0x00B7).Width;
				break;
			default:
				this.LeaderCode  = 0x00;
				this.LeaderWidth = 0;
				break;
		}
	};
	CRunTab.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CRunTab.prototype.GetWidthVisible = function()
	{
		return this.Width;
	};
	CRunTab.prototype.SetWidthVisible = function(WidthVisible)
	{
	};
	CRunTab.prototype.Copy = function()
	{
		return new CRunTab();
	};
	CRunTab.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunTab.prototype.SaveRecalculateObject = function()
	{
		return {
			Width        : this.Width,
			LeaderCode   : this.LeaderCode,
			LeaderWidth  : this.LeaderWidth
		};
	};
	CRunTab.prototype.LoadRecalculateObject = function(RecalcObj)
	{
		this.Width        = RecalcObj.Width;
		this.LeaderCode   = RecalcObj.LeaderCode;
		this.LeaderWidth  = RecalcObj.LeaderWidth;
	};
	CRunTab.prototype.PrepareRecalculateObject = function()
	{
	};
	CRunTab.prototype.GetAutoCorrectFlags = function()
	{
		return AscWord.AUTOCORRECT_FLAGS_ALL;
	};
	CRunTab.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextSpecialTab();
	};
	CRunTab.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	CRunTab.prototype.getBidiType = function()
	{
		return AscBidi.TYPE.PM;
	};
	/**
	 * Absolute position tab (w:ptab): a tab to a position given relative to the margin or the paragraph indent,
	 * not to the tab stops. It behaves like a tab of the matching kind (left/center/right) whose stop is computed
	 * from the page frame when the line is recalculated (see Paragraph.prototype.private_RecalculateGetPTabPos).
	 * @constructor
	 * @extends {AscWord.CRunTab}
	 */
	function CRunPTab()
	{
		CRunTab.call(this);
		
		this.Alignment  = "left";   // left | center | right
		this.RelativeTo = "margin"; // margin | indent
		this.Leader     = "none";   // none | dot | hyphen | underscore | middleDot
	}
	CRunPTab.prototype = Object.create(CRunTab.prototype);
	CRunPTab.prototype.constructor = CRunPTab;
	
	CRunPTab.alignments  = ["left", "center", "right"];
	CRunPTab.relativeTos = ["margin", "indent"];
	CRunPTab.leaders     = ["none", "dot", "hyphen", "underscore", "middleDot"];
	
	CRunPTab.prototype.IsPTab = function()
	{
		return true;
	};
	CRunPTab.prototype.GetAlignment = function()
	{
		return this.Alignment;
	};
	CRunPTab.prototype.GetRelativeTo = function()
	{
		return this.RelativeTo;
	};
	CRunPTab.prototype.GetLeader = function()
	{
		return this.Leader;
	};
	CRunPTab.prototype.SetAlignment = function(value)
	{
		if (-1 !== CRunPTab.alignments.indexOf(value))
			this.Alignment = value;
	};
	CRunPTab.prototype.SetRelativeTo = function(value)
	{
		if (-1 !== CRunPTab.relativeTos.indexOf(value))
			this.RelativeTo = value;
	};
	CRunPTab.prototype.SetLeader = function(value)
	{
		// The inherited SetLeader(leaderType, textPr) is called by the recalculation with an Asc.c_oAscTabLeader
		// value and a text properties object, here a string is a request to change the stored leader
		if (typeof value === "string")
		{
			if (-1 !== CRunPTab.leaders.indexOf(value))
				this.Leader = value;
			return;
		}
		CRunTab.prototype.SetLeader.apply(this, arguments);
	};
	CRunPTab.prototype.GetTabLeaderType = function()
	{
		switch (this.Leader)
		{
			case "dot":        return Asc.c_oAscTabLeader.Dot;
			case "hyphen":     return Asc.c_oAscTabLeader.Hyphen;
			case "underscore": return Asc.c_oAscTabLeader.Underscore;
			case "middleDot":  return Asc.c_oAscTabLeader.MiddleDot;
		}
		return Asc.c_oAscTabLeader.None;
	};
	CRunPTab.prototype.GetTabType = function()
	{
		switch (this.Alignment)
		{
			case "center": return tab_Center;
			case "right":  return tab_Right;
		}
		return tab_Left;
	};
	CRunPTab.prototype.Copy = function()
	{
		let tab = new CRunPTab();
		tab.Alignment  = this.Alignment;
		tab.RelativeTo = this.RelativeTo;
		tab.Leader     = this.Leader;
		return tab;
	};
	CRunPTab.prototype.Write_ToBinary = function(Writer)
	{
		// Long   : para_PTab
		// String : alignment
		// String : relativeTo
		// String : leader
		Writer.WriteLong(para_PTab);
		Writer.WriteString2(this.Alignment);
		Writer.WriteString2(this.RelativeTo);
		Writer.WriteString2(this.Leader);
	};
	CRunPTab.prototype.Read_FromBinary = function(Reader)
	{
		this.SetAlignment(Reader.GetString2());
		this.SetRelativeTo(Reader.GetString2());
		this.SetLeader(Reader.GetString2());
	};
	CRunTab.prototype.IsPTab = function()
	{
		return false;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunTab  = CRunTab;
	window['AscWord'].CRunPTab = CRunPTab;

})(window);
