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
 * Flash Fill (Ctrl+E), as in Microsoft Excel, for the common patterns: the user types one or more
 * examples in a column next to the data, the editor finds a program that builds each example from
 * the cells of its row and fills the empty cells of the column with it.
 *
 * A program is a list of parts: literal text, the whole text of a source cell, a word of a source
 * cell (counted from the start or from the end) or its first letter, each with a case change
 * (as typed, upper, lower, proper). Words are runs of letters and digits. A program is accepted
 * only when it rebuilds every example exactly; the first row is taken as a header when the
 * examples only agree without it.
 */
(function (window) {
	const AscCommonExcel = window["AscCommonExcel"];

	const WORD = /[A-Za-z0-9\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF]+/g;

	function words(text) {
		const res = [];
		WORD.lastIndex = 0;
		let m;
		while ((m = WORD.exec(text))) {
			res.push(m[0]);
		}
		return res;
	}

	function proper(s) {
		return s.toLowerCase().replace(/(^|[^A-Za-z0-9\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF'])([A-Za-z0-9\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF])/g, function (all, before, ch) {
			return before + ch.toUpperCase();
		});
	}

	function applyCase(s, cas) {
		switch (cas) {
			case "upper": return s.toUpperCase();
			case "lower": return s.toLowerCase();
			case "proper": return proper(s);
		}
		return s;
	}

	// the case change that turns `from` into `to`, or null
	function caseOf(from, to) {
		if (from === to) {
			return "same";
		}
		if (from.toUpperCase() === to) {
			return "upper";
		}
		if (from.toLowerCase() === to) {
			return "lower";
		}
		if (proper(from) === to) {
			return "proper";
		}
		return null;
	}

	function evalPart(part, sources, option) {
		if ("lit" === part.t) {
			return part.s;
		}
		const text = sources[part.col];
		if (text == null || text === "") {
			return null;
		}
		if ("all" === part.t) {
			return applyCase(text, part.cas);
		}
		const w = words(text);
		const index = option ? w.length + part.ridx : part.idx;
		if (index < 0 || index >= w.length) {
			return null;
		}
		const word = part.initial ? w[index].charAt(0) : w[index];
		return applyCase(word, part.cas);
	}

	function run(program, options, sources) {
		let out = "";
		for (let i = 0; i < program.length; ++i) {
			const s = evalPart(program[i], sources, options[i]);
			if (null === s) {
				return null;
			}
			out += s;
		}
		return out;
	}

	// The parts that build `output` from `sources`, longest match first
	function synthesize(sources, output) {
		const parts = [];
		const wordsOf = sources.map(function (s) {
			return s ? words(s) : [];
		});
		let pos = 0;
		while (pos < output.length) {
			let best = null;
			for (let col = 0; col < sources.length; ++col) {
				const text = sources[col];
				if (!text) {
					continue;
				}
				const whole = output.substr(pos, text.length);
				const wholeCase = whole.length === text.length ? caseOf(text, whole) : null;
				if (wholeCase && text.length > 1 && (!best || text.length > best.len)) {
					best = {len: text.length, part: {t: "all", col: col, cas: wholeCase}};
				}
				const w = wordsOf[col];
				for (let i = 0; i < w.length; ++i) {
					const piece = output.substr(pos, w[i].length);
					const cas = piece.length === w[i].length ? caseOf(w[i], piece) : null;
					if (cas && (!best || w[i].length > best.len)) {
						best = {len: w[i].length, part: {t: "word", col: col, idx: i, ridx: i - w.length, cas: cas, initial: false}};
					}
				}
				if (!best || best.len < 1) {
					for (let i = 0; i < w.length; ++i) {
						const ch = output.charAt(pos);
						const cas = caseOf(w[i].charAt(0), ch);
						if (cas && /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF]/.test(ch)) {
							// a typed initial gives its case to the other rows
							const initialCase = ch === ch.toUpperCase() ? "upper" : "lower";
							best = {len: 1, part: {t: "word", col: col, idx: i, ridx: i - w.length, cas: initialCase, initial: true}};
							break;
						}
					}
				}
			}
			if (best) {
				parts.push(best.part);
				pos += best.len;
			} else {
				const last = parts[parts.length - 1];
				if (last && "lit" === last.t) {
					last.s += output.charAt(pos);
				} else {
					parts.push({t: "lit", s: output.charAt(pos)});
				}
				pos += 1;
			}
		}
		return parts;
	}

	/**
	 * @param examples {Array<{sources: string[], output: string}>}
	 * @returns {?{parts: Array, options: boolean[]}} the program that rebuilds every example
	 */
	function findProgram(examples) {
		for (let e = 0; e < examples.length; ++e) {
			const parts = synthesize(examples[e].sources, examples[e].output);
			if (!parts.some(function (p) { return "lit" !== p.t; })) {
				continue;
			}
			const variable = [];
			parts.forEach(function (p, i) {
				if ("word" === p.t) {
					variable.push(i);
				}
			});
			const combos = Math.min(1 << variable.length, 256);
			for (let mask = 0; mask < combos; ++mask) {
				// a word taken from the end of the example ("Smith" of "John Smith") is counted from the end first
				const options = parts.map(function () {
					return false;
				});
				variable.forEach(function (index, bit) {
					const p = parts[index];
					const fromEnd = p.idx !== 0 && p.ridx === -1;
					options[index] = fromEnd !== !!(mask & (1 << bit));
				});
				const ok = examples.every(function (ex) {
					return run(parts, options, ex.sources) === ex.output;
				});
				if (ok) {
					return {parts: parts, options: options};
				}
			}
		}
		return null;
	}

	function applyProgram(program, sources) {
		return run(program.parts, program.options, sources);
	}

	/**
	 * Flash Fill of the column of the active cell, inside the current region.
	 * @returns {number} the number of cells filled, -1 when no pattern was found
	 */
	function flashFill(wb) {
		const ws = wb.getWorksheet();
		const model = ws.model;
		const sr = model.selectionRange;
		const active = sr.activeCell;
		const col = active.col;
		const region = model.autoFilters.expandRange(new Asc.Range(col, active.row, col, active.row), true, true, true);
		if (!region || region.c1 === region.c2) {
			return -1;
		}
		const text = function (r, c) {
			return model.getRange3(r, c, r, c).getValueWithFormat() || "";
		};
		const sourceCols = [];
		for (let c = region.c1; c <= region.c2; ++c) {
			if (c !== col) {
				sourceCols.push(c);
			}
		}
		const rows = [];
		for (let r = region.r1; r <= region.r2; ++r) {
			rows.push({row: r, sources: sourceCols.map(function (c) {
				return text(r, c);
			}), output: text(r, col)});
		}
		const examples = rows.filter(function (x) {
			return x.output !== "" && x.sources.some(function (s) {
				return s !== "";
			});
		});
		if (!examples.length) {
			return -1;
		}
		let program = findProgram(examples);
		let first = region.r1;
		if (!program && examples.length > 1 && examples[0].row === region.r1) {
			program = findProgram(examples.slice(1));
			first = region.r1 + 1;
		}
		if (!program) {
			return -1;
		}
		const writes = [];
		rows.forEach(function (x) {
			if (x.row < first || x.output !== "" || !x.sources.some(function (s) { return s !== ""; })) {
				return;
			}
			const value = applyProgram(program, x.sources);
			if (value) {
				writes.push({row: x.row, value: value});
			}
		});
		if (!writes.length) {
			return 0;
		}
		const target = new Asc.Range(col, writes[0].row, col, writes[writes.length - 1].row);
		ws._isLockedCells(target, null, function (success) {
			if (!success) {
				return;
			}
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			writes.forEach(function (w) {
				// Flash Fill writes text: a leading "=" must not become a formula
				const value = "=" === w.value.charAt(0) ? "'" + w.value : w.value;
				model.getRange3(w.row, col, w.row, col).setValue(value);
			});
			AscCommon.History.EndTransaction();
			ws._updateRange(target);
			ws.draw();
			wb.Api.sendEvent("asc_onFlashFill", writes.length);
		});
		return writes.length;
	}

	AscCommonExcel.flashFillFindProgram = findProgram;
	AscCommonExcel.flashFillApplyProgram = applyProgram;

	const api = window["Asc"]["spreadsheet_api"];
	/**
	 * Flash Fill of the column of the active cell.
	 * @returns {number} the number of cells filled, -1 when no pattern was found
	 */
	api.prototype.asc_FlashFill = function () {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit() || !this.wb || this.wb.getCellEditMode()) {
			return -1;
		}
		const res = flashFill(this.wb);
		if (res < 0) {
			this.sendEvent("asc_onFlashFill", -1);
		}
		return res;
	};
	api.prototype["asc_FlashFill"] = api.prototype.asc_FlashFill;
	window["AscCommonExcel"]["flashFillFindProgram"] = findProgram;
	window["AscCommonExcel"]["flashFillApplyProgram"] = applyProgram;
})(window);
