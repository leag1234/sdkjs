// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// A change of the document language restarts the spell check while the answer to the first request is still on its
// way. That stale answer carries the same RecalcId and the same words; it used to mark every word of the new language
// as misspelled (a French Markdown document opened under a French interface was underlined everywhere).
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');

function load() {
	const window = {localStorage: {getItem: () => null, setItem() {}}};
	window.window = window;
	const ctx = vm.createContext({window, console, AscWord: {}, AscCommon: {
		CSpellCheckSettings: function () {}, IsDigit: () => false, IsAbbreviation: () => false
	}});
	vm.runInContext('String.prototype.getUnicodeIterator = function () {' +
		'const cps = Array.from(this).map(c => c.codePointAt(0)); let i = 0;' +
		'return {check: () => i < cps.length, next: () => { i++; }, value: () => cps[i]}; };', ctx);
	ctx.AscWord.SpellMarkStart = function (e) { this.e = e; }; ctx.AscWord.SpellMarkEnd = function (e) { this.e = e; };
	for (const f of ['DocumentSpellChecker.js', 'ParagraphSpellChecker.js', 'ParagraphSpellCheckerElement.js'])
		vm.runInContext(fs.readFileSync(path.join(root, 'word/Editor/SpellChecker', f), 'utf8'), ctx, {filename: f});
	return ctx;
}

function setup() {
	const ctx = load(), requests = [];
	const spelling = new ctx.AscWord.CDocumentSpellChecker();
	const paragraph = {GetId: () => 'p1'};
	ctx.editor = {
		SpellCheckApi: {checkDictionary: () => true, spellCheck: r => requests.push(r)},
		WordControl: {m_oLogicDocument: {Spelling: spelling}},
		asc_IsSpellCheckCurrentWord: () => false
	};
	const checker = new ctx.AscWord.CParagraphSpellChecker(paragraph);
	paragraph.SpellChecker = checker;
	checker.GetCurrentPositionInParagraph = () => null;
	const run = {AddSpellCheckerElement() {}, RemoveSpellCheckerElement() {}};
	const collect = lang => {
		checker.Elements = [];
		for (const w of ['bonjour', 'monde']) checker.Add(run, 0, run, 1, w, lang, 0, 0, 0);
		return checker;
	};
	return {checker, spelling, paragraph, requests, collect};
}

test('an answer for the old language is ignored after a restart', () => {
	const {checker, spelling, paragraph, requests, collect} = setup();
	collect(1033); checker.Check(1);
	const english = requests[0];
	spelling.Reset();                      // language change: RestartSpellCheck
	collect(1036); checker.Check(1);
	assert.equal(requests.length, 2);
	// the English answer arrives first: both words are unknown to the English dictionary
	checker.SpellCheckResponse(1, [false, false], english.usrWords, english.usrLang);
	assert.equal(checker.Elements.length, 2, 'elements still wait for the French answer');
	assert.ok(checker.Elements.every(e => e.IsUndefined()));
	assert.ok(spelling.IsWaitingParagraph(paragraph));
	// the French answer is applied
	const french = requests[1];
	checker.SpellCheckResponse(1, [true, true], french.usrWords, french.usrLang);
	assert.equal(checker.Elements.length, 0, 'correct words leave no spelling mark');
	assert.ok(!spelling.IsWaitingParagraph(paragraph));
});

test('a misspelled word of the new language is still marked', () => {
	const {checker, requests, collect} = setup();
	collect(1036); checker.Check(1);
	const r = requests[0];
	checker.SpellCheckResponse(1, [true, false], r.usrWords, r.usrLang);
	assert.equal(checker.Elements.length, 1);
	assert.equal(checker.Elements[0].GetWord(), 'monde');
	assert.ok(checker.Elements[0].IsWrong());
});

test('an answer without the echoed request keeps the previous behaviour', () => {
	const {checker, collect} = setup();
	collect(1036); checker.Check(1);
	checker.SpellCheckResponse(1, [true, false]);
	assert.equal(checker.Elements.length, 1);
	assert.ok(checker.Elements[0].IsWrong());
});

test('the answer callback passes the request words and languages', () => {
	const src = fs.readFileSync(path.join(root, 'word/api.js'), 'utf8');
	assert.match(src, /SpellCheckResponse\(Obj\["RecalcId"\], Obj\["usrCorrect"\], Obj\["usrWords"\], Obj\["usrLang"\]\)/);
});
