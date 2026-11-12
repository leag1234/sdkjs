/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In particular,
 * Section 7(a) of the GNU AGPL, as amended, applies to this file.
 */

"use strict";

// Word desktop auto-spacing constants measured from Word 365 PDFs (see SPEC.md).
// East Asian -> Latin/digit is 0.25 em; Latin/digit -> East Asian is 0.20 em for
// Times New Roman and 0.22 em for Arial-like fonts.
const AUTOSPACE_EM_E2L = 0.25; // Word desktop PDF measurement (see SPEC.md).
const AUTOSPACE_EM_L2E_TNR = 0.20; // Word desktop PDF measurement (see SPEC.md).
const AUTOSPACE_EM_L2E_OTHER = 0.22; // Word desktop PDF measurement (see SPEC.md).

const AUTOSPACE_PT_TO_MM = 25.4 / 72; // run item widths are in millimetres, font sizes in points

const AUTOSPACE_CLASS = {
	OTHER: 0,
	EAST_ASIAN: 1,
	LATIN: 2,
	DIGIT: 3
};

function isEastAsianCodePoint(codePoint)
{
	if (0x3040 <= codePoint && codePoint <= 0x30FF)
		return true;
	if (0x31F0 <= codePoint && codePoint <= 0x31FF)
		return true;
	if (0x3100 <= codePoint && codePoint <= 0x312F)
		return true;
	if (0x31A0 <= codePoint && codePoint <= 0x31BF)
		return true;
	if (0x3400 <= codePoint && codePoint <= 0x4DBF)
		return true;
	if (0x4E00 <= codePoint && codePoint <= 0x9FFF)
		return true;
	if (0xF900 <= codePoint && codePoint <= 0xFAFF)
		return true;
	if (0x20000 <= codePoint && codePoint <= 0x2FA1F)
		return true;
	if (0xAC00 <= codePoint && codePoint <= 0xD7AF)
		return true;
	if (0x1100 <= codePoint && codePoint <= 0x11FF)
		return true;
	if (0x3130 <= codePoint && codePoint <= 0x318F)
		return true;
	if (0x3005 === codePoint || 0x3006 === codePoint || 0x3007 === codePoint)
		return true;
	return false;
}

function isRtlCodePoint(codePoint)
{
	// Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan, Mandaic, Adlam and the
	// common RTL punctuation ranges. We only need a conservative check here:
	// the classifier is used to decide whether to add a gap, so excluding a
	// few extra RTL punctuation marks is harmless.
	if (0x0590 <= codePoint && codePoint <= 0x08FF)
		return true;
	if (0xFB1D <= codePoint && codePoint <= 0xFDFF)
		return true;
	if (0xFE70 <= codePoint && codePoint <= 0xFEFC)
		return true;
	return false;
}

function isFullWidthOrHalfWidthForm(codePoint)
{
	return (0xFF00 <= codePoint && codePoint <= 0xFFEF);
}

function isLatinLetterCodePoint(codePoint)
{
	// Unicode L* is too broad to enumerate here; we use the ranges that Word
	// measured plus the common letter blocks. The shaper only calls this for
	// characters that are not East Asian, not RTL and not full-width forms.
	if (codePoint < 0x41)
		return false;
	if (codePoint <= 0x5A || (0x61 <= codePoint && codePoint <= 0x7A))
		return true;
	if (0x00C0 <= codePoint && codePoint <= 0x024F)
		return true;
	if (0x0370 <= codePoint && codePoint <= 0x03FF)
		return true;
	if (0x0400 <= codePoint && codePoint <= 0x04FF)
		return true;
	if (0x0530 <= codePoint && codePoint <= 0x058F)
		return true;
	if (0x0900 <= codePoint && codePoint <= 0x0DFF)
		return true;
	if (0x0E00 <= codePoint && codePoint <= 0x0E7F)
		return true;
	if (0x1E00 <= codePoint && codePoint <= 0x1EFF)
		return true;
	if (0x2C60 <= codePoint && codePoint <= 0x2C7F)
		return true;
	return false;
}

function getAutoSpaceClass(codePoint)
{
	if (isEastAsianCodePoint(codePoint))
		return AUTOSPACE_CLASS.EAST_ASIAN;
	if (0x0030 <= codePoint && codePoint <= 0x0039)
		return AUTOSPACE_CLASS.DIGIT;
	if (isRtlCodePoint(codePoint) || isFullWidthOrHalfWidthForm(codePoint))
		return AUTOSPACE_CLASS.OTHER;
	if (isLatinLetterCodePoint(codePoint))
		return AUTOSPACE_CLASS.LATIN;
	return AUTOSPACE_CLASS.OTHER;
}

function getAutoSpaceGapEm(prevClass, currClass, latinFontName)
{
	const isE2L = (AUTOSPACE_CLASS.EAST_ASIAN === prevClass
		&& (AUTOSPACE_CLASS.LATIN === currClass || AUTOSPACE_CLASS.DIGIT === currClass));
	const isL2E = (AUTOSPACE_CLASS.EAST_ASIAN === currClass
		&& (AUTOSPACE_CLASS.LATIN === prevClass || AUTOSPACE_CLASS.DIGIT === prevClass));

	if (isE2L)
		return AUTOSPACE_EM_E2L;
	if (isL2E)
	{
		const fontName = latinFontName || "";
		return (fontName === "Times New Roman" || fontName === "Times")
			? AUTOSPACE_EM_L2E_TNR : AUTOSPACE_EM_L2E_OTHER;
	}
	return 0;
}

// Runs the auto-space pass over a sequence of items. `items` is an array of
// objects exposing GetCodePoint(), IsText(), IsNBSP(), IsPdfText(), GetFontSize()
// and AddAutoSpaceWidth(mm). `latinFontName` is the font of the left item.
// Returns the number of gaps added (for tests).
function applyAutoSpacePass(items, isDE, isDN, latinFontName)
{
	let nAdded = 0;
	let oPrev = null;
	for (let nPos = 0; nPos < items.length; ++nPos)
	{
		let oItem = items[nPos];
		if (!oItem.IsText() || oItem.IsNBSP() || oItem.IsPdfText())
		{
			oPrev = null;
			continue;
		}
		if (oPrev)
		{
			let nPrevClass = getAutoSpaceClass(oPrev.GetCodePoint());
			let nCurrClass = getAutoSpaceClass(oItem.GetCodePoint());
			let isHit = (AUTOSPACE_CLASS.EAST_ASIAN === nPrevClass
				&& ((AUTOSPACE_CLASS.LATIN === nCurrClass && isDE)
					|| (AUTOSPACE_CLASS.DIGIT === nCurrClass && isDN)))
				|| (AUTOSPACE_CLASS.EAST_ASIAN === nCurrClass
					&& ((AUTOSPACE_CLASS.LATIN === nPrevClass && isDE)
						|| (AUTOSPACE_CLASS.DIGIT === nPrevClass && isDN)));
			if (isHit)
			{
				let nGapEm = getAutoSpaceGapEm(nPrevClass, nCurrClass, latinFontName);
				if (0 !== nGapEm)
				{
					oPrev.AddAutoSpaceWidth(oPrev.GetFontSize() * nGapEm * AUTOSPACE_PT_TO_MM);
					++nAdded;
				}
			}
		}
		oPrev = oItem;
	}
	return nAdded;
}

window["AscWord"] = window["AscWord"] || {};
window["AscWord"].AUTOSPACE_CLASS = AUTOSPACE_CLASS;
window["AscWord"].AUTOSPACE_PT_TO_MM = AUTOSPACE_PT_TO_MM;
window["AscWord"].AUTOSPACE_EM_E2L = AUTOSPACE_EM_E2L;
window["AscWord"].AUTOSPACE_EM_L2E_TNR = AUTOSPACE_EM_L2E_TNR;
window["AscWord"].AUTOSPACE_EM_L2E_OTHER = AUTOSPACE_EM_L2E_OTHER;
window["AscWord"].getAutoSpaceClass = getAutoSpaceClass;
window["AscWord"].getAutoSpaceGapEm = getAutoSpaceGapEm;
window["AscWord"].isEastAsianCodePoint = isEastAsianCodePoint;
window["AscWord"].applyAutoSpacePass = applyAutoSpacePass;
