// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../../common/Drawings/Metafile.js'), 'utf8');

function objectFunction(name)
{
	const marker = name + ' : function';
	const start = source.indexOf(marker);
	assert(start >= 0, name);
	const functionStart = source.indexOf('function', start);
	const brace = source.indexOf('{', functionStart);
	let depth = 0;
	for (let i = brace; i < source.length; ++i)
	{
		if (source[i] === '{') ++depth;
		else if (source[i] === '}' && --depth === 0)
			return source.slice(functionStart, i + 1);
	}
	assert.fail('Unterminated ' + name);
}

test('blob-backed images retain the draw-image command prefix in the PDF stream', () =>
{
	const calls = [];
	const imageData = 'data:image/png;base64,AA==';
	const context = vm.createContext({
		CommandType: {ctDrawImageFromFile: 111},
		AscCommon: {g_oDocumentBlobUrls: {getImageBase64: value => value === 'blob:offline-image' ? imageData : null}}
	});
	const drawImage = vm.runInContext('(' + objectFunction('drawImage') + ')', context);
	const target = {Memory: {
		WriteByte: value => calls.push(['byte', value]),
		WriteString2: value => calls.push(['string', value]),
		WriteDouble: value => calls.push(['double', value])
	}};

	drawImage.call(target, 'blob:offline-image', 1, 2, 3, 4, false);
	assert.deepEqual(calls, [
		['byte', 111],
		['string', imageData],
		['double', 1], ['double', 2], ['double', 3], ['double', 4]
	]);
});
