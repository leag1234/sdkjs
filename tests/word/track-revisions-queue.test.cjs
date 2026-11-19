// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Elements without a logic document wait in a queue that only the document editor flushes. The spreadsheet and
// presentation editors must not queue them: chart labels, shapes and cells would stay in memory for the whole
// session (a chart recalculation per cell edit left about a megabyte behind).
// Run: node --test tests/word/track-revisions-queue.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const source = fs.readFileSync(path.resolve(__dirname, '../../word/Editor/revisions/track-revisions-manager.js'), 'utf8');

function load(editorId) {
	const win = {AscWord: {}, AscCommon: {c_oEditorId: {Word: 1, Spreadsheet: 2, Presentation: 4}}, Asc: {editor: {editorId}}};
	win.window = win;
	const ctx = vm.createContext({window: win, AscWord: win.AscWord, AscCommon: win.AscCommon, Asc: win.Asc, console});
	vm.runInContext(source, ctx, {filename: 'track-revisions-manager.js'});
	return win.AscWord;
}
const element = id => ({GetId: () => id, GetLogicDocument: () => null});
const queued = AscWord => !AscWord.CTrackRevisionsManager.prototype.IsAllChecked.call({CheckArray: []});

test('the document editor queues elements that have no logic document yet', () => {
	const AscWord = load(1);
	assert.equal(queued(AscWord), false);
	AscWord.checkElementInRevision(element('a'));
	assert.equal(queued(AscWord), true);
});

test('the spreadsheet and presentation editors do not queue them', () => {
	for (const id of [2, 4]) {
		const AscWord = load(id);
		for (let i = 0; i < 100; ++i)
			AscWord.checkElementInRevision(element('e' + i));
		assert.equal(queued(AscWord), false, 'editor ' + id);
	}
});

test('an element with a document editor logic document is still checked at once', () => {
	const AscWord = load(1), checked = [];
	AscWord.checkElementInRevision({
		GetId: () => 'x',
		GetLogicDocument: () => ({IsDocumentEditor: () => true, GetTrackRevisionsManager: () => ({CheckElement: e => checked.push(e.GetId())})})
	});
	assert.deepEqual(checked, ['x']);
	assert.equal(queued(AscWord), false);
});
