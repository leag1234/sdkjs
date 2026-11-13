// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Tests for the style ids written on save: the ids of the file are kept, new styles get an id made of their name.
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const src = fs.readFileSync(path.join(root, 'word/Editor/Serialize2.js'), 'utf8');

function load() {
	const start = src.indexOf('AscWord.EoComputeStyleIdsForWrite = function');
	const end = src.indexOf('var c_oSerTableTypes = {', start);
	assert(start >= 0 && end > start);
	// the doc comment above the first helper belongs to it; the selector above is not needed here
	const ctx = vm.createContext({console});
	ctx.AscWord = {};
	vm.runInContext(src.slice(start, end), ctx, {filename: 'Serialize2.js'});
	return ctx.AscWord;
}
const W = load();
const plain = v => JSON.parse(JSON.stringify(v));
const styles = list => ({Style: Object.fromEntries(list.map(([id, name, orig]) => [id, {Name: name, EoDocxStyleId: orig}]))});

test('the id of the file is kept', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['724', 'heading 1', 'Titre1'], ['725', 'Titre 1 Car', 'Titre1Car']]));
	assert.deepEqual(plain(map), {'724': 'Titre1', '725': 'Titre1Car'});
});
test('a new style gets an id made of its name', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['724', 'My Style (new)', undefined], ['725', 'Heading 1', undefined]]));
	assert.deepEqual(plain(map), {'724': 'MyStylenew', '725': 'Heading1'});
});
test('a generated id never takes an id of the file, whatever its case', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['1', 'Heading 1', undefined], ['2', 'x', 'heading1'], ['3', 'Heading  1', undefined]]));
	assert.equal(map['2'], 'heading1');
	assert.equal(map['1'], 'Heading12');
	assert.equal(map['3'], 'Heading13');
});
test('the same id read twice is written once, the other style gets a new id', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['1', 'A', 'Same'], ['2', 'B', 'Same']]));
	assert.equal(map['1'], 'Same');
	assert.equal(map['2'], 'B');
});
test('a name without letters or digits, or starting with a digit, still gives a valid id', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['1', '', undefined], ['2', '1st', undefined], ['3', '***', undefined]]));
	assert.deepEqual(Object.values(map).sort(), ['Style', 'Style1st', 'Style2'].sort());
});
test('every id written is unique', () => {
	const map = W.EoComputeStyleIdsForWrite(styles([['1', 'A', undefined], ['2', 'a', undefined], ['3', 'A', undefined], ['4', 'x', 'A']]));
	const values = Object.values(map);
	assert.equal(new Set(values.map(v => v.toLowerCase())).size, values.length);
	assert.equal(map['4'], 'A');
});
test('without a map (clipboard) or for an unknown id the id is written as it is', () => {
	W.EoStyleIdWriteMap = null;
	assert.equal(W.EoStyleIdForWrite('725'), '725');
	W.EoPrepareStyleIdsForWrite(styles([['725', 'Titre 1 Car', 'Titre1Car']]));
	assert.equal(W.EoStyleIdForWrite('725'), 'Titre1Car');
	assert.equal(W.EoStyleIdForWrite('999'), '999');
	assert.equal(W.EoStyleIdForWrite(null), null);
	W.EoStyleIdWriteMap = null;
});
test('a failure while preparing leaves the generated ids', () => {
	W.EoPrepareStyleIdsForWrite(null);
	assert.equal(W.EoStyleIdWriteMap, null);
	assert.equal(W.EoStyleIdForWrite('725'), '725');
});
test('every place that writes a style reference goes through the map', () => {
	for (const needle of ['EoStyleIdForWrite(id)', 'EoStyleIdForWrite(style.BasedOn)', 'EoStyleIdForWrite(style.Next)',
		'EoStyleIdForWrite(style.Link)', 'EoStyleIdForWrite(pPr.PStyle)', 'EoStyleIdForWrite(rPr.RStyle)',
		'EoStyleIdForWrite(styleId)', 'EoStyleIdForWrite(num.NumStyleLink)', 'EoStyleIdForWrite(num.StyleLink)',
		'EoStyleIdForWrite(lvl.PStyle)', 'oNewStyle.EoDocxStyleId = oNewId.id'])
		assert(src.includes(needle), needle);
});
