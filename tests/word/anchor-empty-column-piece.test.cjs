// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Issue 57: a paragraph pushed wholesale into the next column (by a tight-wrap picture that fills the
// column) keeps an empty piece in the previous column. Paragraph-relative anchors must use the first
// piece that holds lines, not that empty piece (Word anchors the picture at the column where the text is).
// Run: node --test tests/word/anchor-empty-column-piece.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');

for (const [file, who] of [['word/Editor/Run.js', 'Para'], ['word/Editor/Paragraph.js', 'this']])
	test(file + ' skips empty paragraph pieces when choosing the anchor piece', () => {
		const src = read(file);
		const i = src.indexOf('_CurPage = CurPage - ColumnAbs;');
		assert(i >= 0);
		const near = src.slice(i, i + 400);
		assert(near.includes('while (_CurPage < CurPage && ' + who + '.Pages[_CurPage].EndLine < ' + who + '.Pages[_CurPage].StartLine)'));
		assert(near.includes('++_CurPage;'));
	});
