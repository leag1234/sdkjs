/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

(function()
{
	/**
	 * East Asian document grid of a section (w:docGrid)
	 * Type: 0 default (no grid), 1 lines, 2 linesAndChars, 3 snapToChars
	 * LinePitch is in twips, CharSpace in 1/4096 pt (added to the default font size)
	 *
	 * @param {number} [nType]
	 * @param {number} [nLinePitch]
	 * @param {number} [nCharSpace]
	 * @constructor
	 */
	function SectionDocGrid(nType, nLinePitch, nCharSpace)
	{
		this.Type      = undefined !== nType ? nType : 0;
		this.LinePitch = undefined !== nLinePitch ? nLinePitch : undefined;
		this.CharSpace = undefined !== nCharSpace ? nCharSpace : undefined;
	}
	SectionDocGrid.prototype.Copy = function()
	{
		return new SectionDocGrid(this.Type, this.LinePitch, this.CharSpace);
	};
	SectionDocGrid.prototype.IsEqual = function(oOther)
	{
		return !!oOther && this.Type === oOther.Type && this.LinePitch === oOther.LinePitch && this.CharSpace === oOther.CharSpace;
	};
	/** @returns {boolean} true when lines snap to the grid */
	SectionDocGrid.prototype.IsLines = function()
	{
		return this.Type >= 1 && this.LinePitch > 0;
	};
	/** @returns {boolean} true when East Asian characters snap to a character pitch */
	SectionDocGrid.prototype.IsChars = function()
	{
		return this.Type === 2 || this.Type === 3;
	};
	/** @returns {number} line pitch in mm, or 0 when the grid does not snap lines */
	SectionDocGrid.prototype.GetLinePitchMM = function()
	{
		return this.IsLines() ? this.LinePitch * 25.4 / 1440 : 0;
	};
	SectionDocGrid.prototype.Write_ToBinary = function(oWriter)
	{
		let nFlags    = 0;
		let nStartPos = oWriter.GetCurPosition();
		oWriter.Skip(4);
		oWriter.WriteLong(this.Type);
		if (undefined !== this.LinePitch)
		{
			oWriter.WriteLong(this.LinePitch);
			nFlags |= 1;
		}
		if (undefined !== this.CharSpace)
		{
			oWriter.WriteLong(this.CharSpace);
			nFlags |= 2;
		}
		let nEndPos = oWriter.GetCurPosition();
		oWriter.Seek(nStartPos);
		oWriter.WriteLong(nFlags);
		oWriter.Seek(nEndPos);
	};
	SectionDocGrid.prototype.Read_FromBinary = function(oReader)
	{
		let nFlags     = oReader.GetLong();
		this.Type      = oReader.GetLong();
		this.LinePitch = (nFlags & 1) ? oReader.GetLong() : undefined;
		this.CharSpace = (nFlags & 2) ? oReader.GetLong() : undefined;
	};

	//--------------------------------------------------------export----------------------------------------------------
	AscWord.SectionDocGrid = SectionDocGrid;
})();
