/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
'use strict';
// Word repeats the header rows (w:tblHeader) of a floating table (w:tblpPr) on the
// continuation pages exactly like those of an inline table. The recalculation used to
// restrict the header-row page logic to inline tables.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const src = fs.readFileSync(path.join(__dirname, '../../word/Editor/Table/TableRecalculate.js'), 'utf8');
const m = /\n\s*if \((this\.HeaderInfo\.Count > 0 && this\.HeaderInfo\.PageIndex != -1 && CurPage > this\.HeaderInfo\.PageIndex[^\n]*)\)\n/.exec(src);

test('the header-row repeat condition is found', () => {
	assert(m, 'header repeat condition present in private_RecalculatePageY');
});

function repeats(inline, count, pageIndex, curPage) {
	const ctx = {HeaderInfo: {Count: count, PageIndex: pageIndex}, CurPage: curPage, IsInline: () => inline};
	return vm.runInNewContext('(function(){ return ' + m[1].replace(/this\./g, 'ctx.').replace(/CurPage/g, 'ctx.CurPage') + '; })()', {ctx});
}

test('floating and inline tables both repeat header rows on later pages', () => {
	for (const inline of [true, false]) {
		assert.equal(repeats(inline, 1, 0, 1), true);
		assert.equal(repeats(inline, 1, 0, 0), false, 'never on the first page');
		assert.equal(repeats(inline, 0, 0, 1), false, 'no header rows');
		assert.equal(repeats(inline, 1, -1, 1), false, 'header never fitted');
	}
});
