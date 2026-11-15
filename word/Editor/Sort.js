// SPDX-License-Identifier: AGPL-3.0-only
/*
 * Word-style Sort (Home > Paragraph > Sort and Table Layout > Data > Sort).
 *
 * The first part is a pure engine (keys, collation, stable order, minimal move plan).
 * The second part adds CDocument.prototype.GetSortInfo / SortSelection, which move whole
 * paragraphs or table rows inside one history point (nothing is retyped, so every
 * paragraph and row keeps its formatting).
 */

"use strict";

(function(window)
{
	const SORT_TEXT   = 0;
	const SORT_NUMBER = 1;
	const SORT_DATE   = 2;

	const SEP_TAB   = 'tab';
	const SEP_COMMA = 'comma';
	const SEP_OTHER = 'other';

	const MAX_KEYS = 3;

	//------------------------------------------------------------------------------------------------------------------
	// Pure engine
	//------------------------------------------------------------------------------------------------------------------
	function GetSeparator(sKind, sChar)
	{
		if (SEP_COMMA === sKind)
			return ',';
		if (SEP_OTHER === sKind)
			return sChar ? String(sChar) : '';
		return '\t';
	}

	/**
	 * Splits a paragraph text into fields. Without a usable separator the whole text is the only field.
	 * @returns {string[]}
	 */
	function SplitFields(sText, sKind, sChar)
	{
		let sSep = GetSeparator(sKind, sChar);
		if (!sSep)
			return [sText];
		return sText.split(sSep);
	}

	function GetDecimalSymbol(sLang)
	{
		try
		{
			let arrParts = new Intl.NumberFormat(sLang || 'en').formatToParts(1.1);
			for (let nIndex = 0; nIndex < arrParts.length; ++nIndex)
			{
				if ('decimal' === arrParts[nIndex].type)
					return arrParts[nIndex].value;
			}
		}
		catch (e)
		{
		}
		return '.';
	}

	/**
	 * Takes the first number found in the text (Word ignores the other characters).
	 * @returns {?number} null when there is no number
	 */
	function ParseNumber(sText, sDecimal)
	{
		let oMatch = /([-−]?)\s*\(?(\d[\d.,   ]*\d|\d)/.exec(sText);
		if (!oMatch)
			return null;

		let sNum = oMatch[2].replace(/[   ]/g, '');
		let isNeg = !!oMatch[1];
		let nDot = sNum.indexOf('.'), nComma = sNum.indexOf(',');
		let nLastDot = sNum.lastIndexOf('.'), nLastComma = sNum.lastIndexOf(',');
		let sDec = '';

		if (nDot >= 0 && nComma >= 0)
			sDec = nLastDot > nLastComma ? '.' : ',';
		else if (nDot >= 0 || nComma >= 0)
		{
			let sCh = nDot >= 0 ? '.' : ',';
			let isMultiple = sNum.indexOf(sCh) !== sNum.lastIndexOf(sCh);
			if (!isMultiple && sCh === sDecimal)
				sDec = sCh;
			else if (!isMultiple && sNum.length - sNum.indexOf(sCh) - 1 !== 3)
				sDec = sCh;
		}

		let sInt = sNum, sFrac = '';
		if (sDec)
		{
			let nPos = sNum.lastIndexOf(sDec);
			sInt  = sNum.substring(0, nPos);
			sFrac = sNum.substring(nPos + 1);
		}
		sInt = sInt.replace(/[.,]/g, '');

		let nValue = parseFloat(sInt + (sFrac ? '.' + sFrac.replace(/[.,]/g, '') : ''));
		if (!isFinite(nValue))
			return null;
		return isNeg ? -nValue : nValue;
	}

	/**
	 * True when the language writes the month before the day in numeric dates
	 */
	function IsMonthFirst(sLang)
	{
		try
		{
			let arrParts = new Intl.DateTimeFormat(sLang || 'en', {timeZone : 'UTC'}).formatToParts(new Date(Date.UTC(2001, 10, 23)));
			for (let nIndex = 0; nIndex < arrParts.length; ++nIndex)
			{
				if ('month' === arrParts[nIndex].type)
					return true;
				if ('day' === arrParts[nIndex].type)
					return false;
			}
		}
		catch (e)
		{
		}
		return true;
	}

	/**
	 * @returns {?number} a UTC timestamp, null when the text is not a date
	 */
	function ParseDate(sText, sLang)
	{
		let sValue = sText.trim();
		if (!sValue)
			return null;

		let oMatch = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?!\d)/.exec(sValue);
		if (oMatch)
			return MakeDate(+oMatch[1], +oMatch[2], +oMatch[3]);

		oMatch = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2}|\d{4})(?!\d)/.exec(sValue);
		if (oMatch)
		{
			let nFirst = +oMatch[1], nSecond = +oMatch[2], nYear = +oMatch[3];
			if (oMatch[3].length <= 2)
				nYear += nYear < 30 ? 2000 : 1900;

			let isMonthFirst = IsMonthFirst(sLang);
			if (nFirst > 12)
				isMonthFirst = false;
			else if (nSecond > 12)
				isMonthFirst = true;

			return isMonthFirst ? MakeDate(nYear, nFirst, nSecond) : MakeDate(nYear, nSecond, nFirst);
		}

		let nParsed = Date.parse(sValue);
		return isFinite(nParsed) ? nParsed : null;
	}

	function MakeDate(nYear, nMonth, nDay)
	{
		if (nMonth < 1 || nMonth > 12 || nDay < 1 || nDay > 31)
			return null;
		return Date.UTC(nYear, nMonth - 1, nDay);
	}

	function CreateCollator(sLang, isCaseSensitive)
	{
		let oOptions = {sensitivity : isCaseSensitive ? 'variant' : 'accent', caseFirst : 'lower'};
		try
		{
			return new Intl.Collator(sLang || undefined, oOptions);
		}
		catch (e)
		{
			return new Intl.Collator(undefined, oOptions);
		}
	}

	/**
	 * Computes the new order. Stable, so equal items keep their relative order.
	 * Empty values (no text, no number, no date) come first in ascending order, like in Word.
	 * @param {number} nCount
	 * @param {function(number, number): string} fGetField returns the text of the field (0 - whole item) of an item
	 * @param {{field: number, type: number, ascending: boolean}[]} arrKeys
	 * @param {{lang: string, caseSensitive: boolean}} oOptions
	 * @returns {number[]} order[newPos] = oldPos
	 */
	function ComputeOrder(nCount, fGetField, arrKeys, oOptions)
	{
		let sLang          = oOptions && oOptions.lang ? oOptions.lang : '';
		let isCaseSensitive = !!(oOptions && oOptions.caseSensitive);
		let oCollator      = CreateCollator(sLang, isCaseSensitive);
		let sDecimal       = GetDecimalSymbol(sLang);

		let arrUsedKeys = arrKeys.slice(0, MAX_KEYS);
		let arrValues   = [];
		for (let nIndex = 0; nIndex < nCount; ++nIndex)
		{
			let arrItem = [];
			for (let nKey = 0; nKey < arrUsedKeys.length; ++nKey)
			{
				let oKey  = arrUsedKeys[nKey];
				let sText = fGetField(nIndex, oKey.field);
				if (SORT_NUMBER === oKey.type)
					arrItem.push(ParseNumber(sText, sDecimal));
				else if (SORT_DATE === oKey.type)
					arrItem.push(ParseDate(sText, sLang));
				else
					arrItem.push(sText);
			}
			arrValues.push(arrItem);
		}

		let arrOrder = [];
		for (let nIndex = 0; nIndex < nCount; ++nIndex)
			arrOrder.push(nIndex);

		arrOrder.sort(function(nA, nB)
		{
			for (let nKey = 0; nKey < arrUsedKeys.length; ++nKey)
			{
				let oKey = arrUsedKeys[nKey];
				let vA = arrValues[nA][nKey], vB = arrValues[nB][nKey];
				let nResult;
				if (SORT_TEXT === oKey.type)
				{
					nResult = oCollator.compare(vA, vB);
				}
				else if (null === vA || null === vB)
				{
					nResult = (null === vA ? 0 : 1) - (null === vB ? 0 : 1);
				}
				else
				{
					nResult = vA < vB ? -1 : (vA > vB ? 1 : 0);
				}

				if (0 !== nResult)
					return oKey.ascending ? nResult : -nResult;
			}
			return nA - nB;
		});

		return arrOrder;
	}

	/**
	 * Produces the moves that turn the current sequence into the wanted order,
	 * moving as few items as the simple placement allows.
	 * @param {Array} arrItems
	 * @param {number[]} arrOrder arrOrder[newPos] = oldPos
	 * @returns {{from: number, to: number}[]}
	 */
	function PlanMoves(arrItems, arrOrder)
	{
		let arrCurrent = arrItems.slice();
		let arrMoves   = [];
		for (let nPos = 0; nPos < arrOrder.length; ++nPos)
		{
			let oWanted = arrItems[arrOrder[nPos]];
			if (arrCurrent[nPos] === oWanted)
				continue;

			let nFrom = arrCurrent.indexOf(oWanted, nPos);
			arrCurrent.splice(nFrom, 1);
			arrCurrent.splice(nPos, 0, oWanted);
			arrMoves.push({from : nFrom, to : nPos});
		}
		return arrMoves;
	}

	const SortEngine = {
		SORT_TEXT        : SORT_TEXT,
		SORT_NUMBER      : SORT_NUMBER,
		SORT_DATE        : SORT_DATE,
		SEP_TAB          : SEP_TAB,
		SEP_COMMA        : SEP_COMMA,
		SEP_OTHER        : SEP_OTHER,
		MAX_KEYS         : MAX_KEYS,
		SplitFields      : SplitFields,
		ParseNumber      : ParseNumber,
		ParseDate        : ParseDate,
		GetDecimalSymbol : GetDecimalSymbol,
		ComputeOrder     : ComputeOrder,
		PlanMoves        : PlanMoves
	};

	//------------------------------------------------------------------------------------------------------------------
	// Document layer
	//------------------------------------------------------------------------------------------------------------------
	const ERR_NONE    = 'none';    // fewer than two items to sort
	const ERR_MERGED  = 'merged';  // merged cells or grid gaps
	const ERR_NESTED  = 'nested';  // nested tables
	const ERR_SECTION = 'section'; // section break inside the paragraphs
	const ERR_MIXED   = 'mixed';   // selection spans several containers
	const ERR_LOCKED  = 'locked';  // protected or locked content, read-only
	const ERR_OK      = 'ok';

	function GetParagraphText(oParagraph, sTab)
	{
		return oParagraph.GetText({
			Numbering        : false,
			Math             : false,
			ParaSeparator    : '',
			TabSymbol        : sTab,
			NewLineSeparator : ' '
		});
	}

	function GetCellText(oCell)
	{
		let arrContent = oCell.GetContent().Content;
		let arrText    = [];
		for (let nIndex = 0; nIndex < arrContent.length; ++nIndex)
		{
			if (arrContent[nIndex].IsParagraph && arrContent[nIndex].IsParagraph())
				arrText.push(GetParagraphText(arrContent[nIndex], ' '));
		}
		return arrText.join(' ');
	}

	/**
	 * Finds what the Sort command works on.
	 * @returns {{mode: string, error: ?string, ...}}
	 */
	function FindTarget(oDocument, isHeader)
	{
		let oTable = oDocument.GetCurrentTable();
		if (oTable)
			return FindTableTarget(oTable, isHeader);

		return FindParagraphTarget(oDocument);
	}

	function FindParagraphTarget(oDocument)
	{
		let oResult = {mode : 'paragraphs', error : null, container : null, start : 0, count : 0, items : []};
		if (!oDocument.IsSelectionUse() || !oDocument.IsTextSelectionUse())
		{
			oResult.error = ERR_NONE;
			return oResult;
		}

		let arrParagraphs = oDocument.GetSelectedParagraphs();
		if (!arrParagraphs || arrParagraphs.length < 2)
		{
			oResult.error = ERR_NONE;
			return oResult;
		}

		let oParent = arrParagraphs[0].GetParent();
		let nStart  = oParent && oParent.Content ? oParent.Content.indexOf(arrParagraphs[0]) : -1;
		if (nStart < 0)
		{
			oResult.error = ERR_MIXED;
			return oResult;
		}

		for (let nIndex = 0; nIndex < arrParagraphs.length; ++nIndex)
		{
			if (oParent.Content[nStart + nIndex] !== arrParagraphs[nIndex])
			{
				oResult.error = ERR_MIXED;
				return oResult;
			}

			if (arrParagraphs[nIndex].Get_SectionPr && arrParagraphs[nIndex].Get_SectionPr())
			{
				oResult.error = ERR_SECTION;
				return oResult;
			}
		}

		oResult.container = oParent;
		oResult.start     = nStart;
		oResult.count     = arrParagraphs.length;
		oResult.items     = arrParagraphs;
		return oResult;
	}

	function FindTableTarget(oTable, isHeader)
	{
		let oResult = {mode : 'table', error : null, table : oTable, start : 0, count : 0, items : [], columns : 0, header : false};

		let nRowsCount = oTable.GetRowsCount();
		let nStart = 0, nEnd = nRowsCount - 1;
		if (oTable.IsCellSelection() && !oTable.ApplyToAll && oTable.Selection && oTable.Selection.Data && oTable.Selection.Data.length)
		{
			let nMin = nRowsCount, nMax = -1;
			for (let nIndex = 0; nIndex < oTable.Selection.Data.length; ++nIndex)
			{
				let nRow = oTable.Selection.Data[nIndex].Row;
				nMin = Math.min(nMin, nRow);
				nMax = Math.max(nMax, nRow);
			}
			if (nMax - nMin >= 1)
			{
				nStart = nMin;
				nEnd   = nMax;
			}
		}

		oResult.start = nStart;
		oResult.count = nEnd - nStart + 1;
		oResult.header = undefined === isHeader ? !!oTable.GetRow(nStart).IsHeader() : !!isHeader;

		for (let nRow = nStart; nRow <= nEnd; ++nRow)
			oResult.items.push(oTable.GetRow(nRow));

		if (oResult.count < 2)
		{
			oResult.error = ERR_NONE;
			return oResult;
		}

		// The first row may be a merged title: it is only checked when it takes part in the sort,
		// but the column count comes from the rows below it.
		let nColumns = oResult.items[oResult.items.length - 1].GetCellsCount();
		oResult.columns = nColumns;

		for (let nIndex = 0; nIndex < oResult.items.length; ++nIndex)
		{
			let oRow = oResult.items[nIndex];
			if (0 === nIndex && oResult.header && oRow.GetCellsCount() !== nColumns)
				continue;

			if (oRow.GetCellsCount() !== nColumns || oRow.GetBefore().GridBefore > 0 || oRow.GetAfter().GridAfter > 0)
			{
				oResult.error = ERR_MERGED;
				return oResult;
			}

			for (let nCell = 0; nCell < nColumns; ++nCell)
			{
				let oCell = oRow.GetCell(nCell);
				if (oCell.GetGridSpan() > 1)
				{
					if (0 === nIndex && oResult.header)
						continue;
					oResult.error = ERR_MERGED;
					return oResult;
				}

				if (CellHasVerticalMerge(oTable, oCell, nStart + nIndex, nCell))
				{
					oResult.error = ERR_MERGED;
					return oResult;
				}

				let arrContent = oCell.GetContent().Content;
				for (let nElement = 0; nElement < arrContent.length; ++nElement)
				{
					if (arrContent[nElement].IsTable && arrContent[nElement].IsTable())
					{
						oResult.error = ERR_NESTED;
						return oResult;
					}
				}
			}
		}

		return oResult;
	}

	function CellHasVerticalMerge(oTable, oCell, nRow, nCell)
	{
		if (typeof vmerge_Continue === 'undefined')
			return false;

		if (vmerge_Continue === oCell.GetVMerge())
			return true;

		if (nRow + 1 < oTable.GetRowsCount())
		{
			let oNextRow = oTable.GetRow(nRow + 1);
			if (nCell < oNextRow.GetCellsCount() && vmerge_Continue === oNextRow.GetCell(nCell).GetVMerge())
				return true;
		}

		return false;
	}

	function CountFields(oTarget, sKind, sChar)
	{
		let nMax = 1;
		for (let nIndex = 0; nIndex < oTarget.items.length; ++nIndex)
		{
			let sText = GetParagraphText(oTarget.items[nIndex], '\t');
			nMax = Math.max(nMax, SplitFields(sText, sKind, sChar).length);
		}
		return nMax;
	}

	/**
	 * Describes the current selection for the Sort dialog and the toolbar.
	 * @param {{separator: string, separatorChar: string}} [oProps]
	 */
	CDocument.prototype.GetSortInfo = function(oProps)
	{
		let oTarget = FindTarget(this, oProps ? oProps["header"] : undefined);
		let oInfo   = {};
		oInfo['mode']    = oTarget.mode;
		oInfo['error']   = oTarget.error;
		oInfo['count']   = oTarget.count;
		oInfo['columns'] = oTarget.columns || 0;
		oInfo['fields']  = 1;
		oInfo['header']  = !!oTarget.header;

		if ('paragraphs' === oTarget.mode && !oTarget.error)
			oInfo['fields'] = CountFields(oTarget, oProps ? oProps['separator'] : SEP_TAB, oProps ? oProps['separatorChar'] : '');

		return oInfo;
	};

	/**
	 * Sorts the selected paragraphs or table rows.
	 * @param {Object} oProps keys: [{field, type, ascending}], header, caseSensitive, separator, separatorChar, lang
	 * @returns {string} 'ok' or one of the error codes
	 */
	CDocument.prototype.SortSelection = function(oProps)
	{
		let oTarget = FindTarget(this, !!oProps['header']);
		if (oTarget.error)
			return oTarget.error;

		let isTable = 'table' === oTarget.mode;
		let nSkip   = oProps['header'] ? 1 : 0;
		let nCount  = oTarget.count - nSkip;
		if (nCount < 2)
			return ERR_NONE;

		let arrKeys = [];
		let arrInKeys = oProps['keys'] || [];
		for (let nIndex = 0; nIndex < arrInKeys.length && arrKeys.length < MAX_KEYS; ++nIndex)
		{
			let oKey = arrInKeys[nIndex];
			if (!oKey || oKey['field'] < 0)
				continue;
			arrKeys.push({field : oKey['field'] | 0, type : oKey['type'] | 0, ascending : false !== oKey['ascending']});
		}
		if (!arrKeys.length)
			return ERR_NONE;

		let arrItems = oTarget.items.slice(nSkip);
		let sKind    = oProps['separator'] || SEP_TAB, sChar = oProps['separatorChar'] || '';

		let fGetField;
		if (isTable)
		{
			// field is the 1-based column number
			fGetField = function(nIndex, nField)
			{
				let oCell = arrItems[nIndex].GetCell(Math.min(Math.max(nField, 1), arrItems[nIndex].GetCellsCount()) - 1);
				return GetCellText(oCell);
			};
		}
		else
		{
			// field 0 is the whole paragraph, field N is the N-th separated field (1-based)
			fGetField = function(nIndex, nField)
			{
				let sText = GetParagraphText(arrItems[nIndex], '\t');
				if (nField <= 0)
					return sText;
				let arrFields = SplitFields(sText, sKind, sChar);
				return nField <= arrFields.length ? arrFields[nField - 1] : '';
			};
		}

		let arrOrder = ComputeOrder(nCount, fGetField, arrKeys, {lang : oProps['lang'], caseSensitive : !!oProps['caseSensitive']});
		let arrMoves = PlanMoves(arrItems, arrOrder);
		if (!arrMoves.length)
			return ERR_OK;

		let nFirst = oTarget.start + nSkip;
		let oState = this.SaveDocumentState(false);

		if (this.IsSelectionLocked(isTable ? AscCommon.changestype_Table_Properties : AscCommon.changestype_Remove))
		{
			this.LoadDocumentState(oState);
			this.UpdateSelection();
			return ERR_LOCKED;
		}

		this.StartAction(AscDFH.historydescription_Document_SortSelection);

		let isPrevented = this.PreventPreDelete;
		this.PreventPreDelete = true;
		for (let nMove = 0; nMove < arrMoves.length; ++nMove)
		{
			let nFrom = nFirst + arrMoves[nMove].from, nTo = nFirst + arrMoves[nMove].to;
			if (isTable)
			{
				let oTable = oTarget.table;
				let oRow   = oTable.GetRow(nFrom);
				oTable.private_RemoveRow(nFrom);
				oTable.private_AddRow(nTo, oRow.GetCellsCount(), true, oRow);
			}
			else
			{
				let oContainer = oTarget.container;
				let oParagraph = oContainer.Content[nFrom];
				oContainer.Remove_FromContent(nFrom, 1, false);
				oContainer.Add_ToContent(nTo, oParagraph, false);
			}
		}
		this.PreventPreDelete = isPrevented;

		if (isTable)
			oTarget.table.Recalc_CompiledPr2();

		this.LoadDocumentState(oState);
		if (!isTable && oTarget.container === this)
			this.private_SelectSortedParagraphs(oTarget.start, oTarget.count);

		this.UpdateSelection();
		this.Recalculate();
		this.FinalizeAction();
		return ERR_OK;
	};

	/**
	 * Selects again the paragraphs that were sorted (top level of the document)
	 */
	CDocument.prototype.private_SelectSortedParagraphs = function(nStart, nCount)
	{
		this.RemoveSelection();

		let nEnd = nStart + nCount - 1;
		this.Selection.Use      = true;
		this.Selection.Start    = false;
		this.Selection.Flag     = selectionflag_Common;
		this.Selection.StartPos = nStart;
		this.Selection.EndPos   = nEnd;
		this.CurPos.ContentPos  = nEnd;
		for (let nIndex = nStart; nIndex <= nEnd; ++nIndex)
			this.Content[nIndex].SelectAll(1);
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].SortEngine = SortEngine;

})(window);
