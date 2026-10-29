// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../../common/libfont/map.js'), 'utf8');

function assignedFunction(name)
{
	const marker = 'this.' + name + ' = function';
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

test('a known symbol face missing from the runtime inventory falls back to the registered default', () =>
{
	// Public repro: GOV.UK Theory of Change template (OGL v3), SHA-256 b0f13034…,
	// declares Marlett in word/fontTable.xml. The dictionary can select Marlett even
	// when the packaged runtime has no matching font-info entry.
	const fallback = {m_wsFontName: 'Carlito'};
	const selected = {m_wsFontName: 'Marlett'};
	const app = {
		FontPickerMap: {},
		DefaultIndex: 0,
		g_fontSelections: {List: [fallback]},
		g_fontDictionary: {FD_UnknownFontSubstitute: 'Calibri'},
		GetSubstituteName: name => name,
		GetInstalledSubstitute: name => name === 'Calibri' ? 'Carlito' : null,
		GetFontIndex: select => select.wsName === 'Carlito' ? fallback : selected
	};
	const context = vm.createContext({AscFonts: {g_map_font_index: {Carlito: 0}}, window: {}, CFontSelectFormat: function() {}});
	app.GetFontFileWeb = vm.runInContext('(' + assignedFunction('GetFontFileWeb') + ')', context);
	const actual = app.GetFontFileWeb('Marlett', 0);
	assert.strictEqual(actual, fallback);
	assert.strictEqual(app.FontPickerMap.Marlett, fallback);
});
