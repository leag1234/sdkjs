// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// Fonts embedded in DOCX (w:embedRegular & co., word/fonts/*.odttf): the font table items written by the converter
// (EmbedRegular..EmbedBoldItalic of c_oSerFontTableTypes), the reader in Serialize2.js and the registration of the
// faces for the open document in the font engine (map.js, GlobalLoaders.js).
// Run: node --test tests/word/embedded-fonts.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm'), cp = require('node:child_process');
const root = path.resolve(__dirname, '../..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const serial = read('word/Editor/Serialize2.js');
const mapSrc = read('common/libfont/map.js');
const loaderSrc = read('common/GlobalLoaders.js');

function slice(source, from, to) {
	const start = source.indexOf(from);
	assert(start >= 0, from);
	const end = to === undefined ? source.length : source.indexOf(to, start + from.length);
	assert(end > start, to);
	return source.slice(start, end);
}
function assignedFunction(source, name) {
	const marker = 'this.' + name + ' = function';
	const start = source.indexOf(marker);
	assert(start >= 0, name);
	const functionStart = source.indexOf('function', start);
	const brace = source.indexOf('{', functionStart);
	let depth = 0;
	for (let i = brace; i < source.length; ++i) {
		if (source[i] === '{') ++depth;
		else if (source[i] === '}' && --depth === 0)
			return source.slice(functionStart, i + 1);
	}
	assert.fail('Unterminated ' + name);
}
const eq = (a, b, m) => assert.equal(JSON.stringify(a), JSON.stringify(b), m);
function committed(file) {
	return cp.execFileSync('git', ['show', 'HEAD:' + file], {cwd: root, maxBuffer: 1 << 28}).toString('utf8');
}

// ---- common serialization code (Binary_CommonReader, FT_Stream2) ----
const win = {AscCommon: {}, Asc: {}};
const commonCtx = vm.createContext({
	window: win, AscCommon: win.AscCommon, Asc: win.Asc, console, Uint8Array, TextDecoder,
	AscFonts: {FT_Common: {UintToInt: v => v | 0, IntToUInt: v => v >>> 0}}
});
vm.runInContext(read('common/SerializeCommonWordExcel.js'), commonCtx, {filename: 'SerializeCommonWordExcel.js'});
const AC = win.AscCommon;

// ---- helpers building the binary font table like BinaryWriterD.cpp does ----
function utf16(s) {
	const a = [];
	for (let i = 0; i < s.length; ++i) a.push(s.charCodeAt(i) & 0xFF, s.charCodeAt(i) >> 8);
	return a;
}
const long = v => [v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF];
// item: type byte, length (long), payload; the font table is read with the same read-table layout: long length + items
const item = (type, payload) => [type].concat(long(payload.length), payload);
const strItem = (type, s) => item(type, utf16(s));
function fontTable(fonts) {
	const body = [].concat(...fonts.map(f => item(0, [].concat(...f))));
	return Uint8Array.from(long(body.length).concat(body));
}

function loadReader(source) {
	const ctx = vm.createContext({
		AscCommon: AC, Asc: win.Asc, console, c_oSerConstants: AC.c_oSerConstants, c_oSerPropLenType: AC.c_oSerPropLenType,
		Binary_CommonReader: AC.Binary_CommonReader
	});
	const code = slice(source, 'var c_oSerFontTableTypes = {', 'var c_oSerSigTypes')
		+ slice(source, 'function Binary_FontTableReader(', 'function Binary_SettingsTableReader(')
		+ ';this.Types = c_oSerFontTableTypes; this.Reader = Binary_FontTableReader;';
	vm.runInContext(code, ctx, {filename: 'Serialize2.js'});
	return ctx;
}
const cur = loadReader(serial);
function readTable(ctx, bytes) {
	const result = {fontAltNames: {}, fontFamilies: {}, embeddedFonts: []};
	const stream = new AC.FT_Stream2(bytes, bytes.length);
	const res = new ctx.Reader(result, stream).Read();
	return {res, result};
}

test('item numbers of sdkjs equal core BinReaderWriterDefines.h and old numbers are not changed', () => {
	eq(Object.assign({}, cur.Types), {
		Font: 0, Name: 1, AltName: 2, Family: 3, EmbedRegular: 4, EmbedBold: 5, EmbedItalic: 6, EmbedBoldItalic: 7,
		EmbedKeyRegular: 8, EmbedKeyBold: 9, EmbedKeyItalic: 10, EmbedKeyBoldItalic: 11, EmbedSubsetted: 12
	});
	const core = path.resolve(root, '../core/OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h');
	if (!fs.existsSync(core))
		return;
	const header = slice(fs.readFileSync(core, 'utf8'), 'namespace c_oSerFontTableTypes', '};}');
	for (const [k, v] of Object.entries(cur.Types)) {
		const m = new RegExp('\\b' + k + '\\s*=\\s*(\\d+)').exec(header);
		assert(m, 'core has ' + k);
		assert.equal(+m[1], v, k);
	}
});

test('the font table reader collects the embedded faces of a font', () => {
	const bytes = fontTable([
		[strItem(1, 'Zz Embedded Test'), strItem(3, 'swiss'), strItem(4, 'fonts/Zz_Embedded_Test_regular.ttf'),
			strItem(5, 'fonts/Zz_Embedded_Test_bold.otf'), strItem(7, 'fonts/Zz_Embedded_Test_bolditalic.ttf')],
		[strItem(1, 'Arial'), strItem(2, 'Arimo')]
	]);
	const {res, result} = readTable(cur, bytes);
	assert.equal(res, AC.c_oSerConstants.ReadOk);
	eq(result.embeddedFonts, [
		{name: 'Zz Embedded Test', slot: 0, path: 'fonts/Zz_Embedded_Test_regular.ttf', key: '', subsetted: false},
		{name: 'Zz Embedded Test', slot: 1, path: 'fonts/Zz_Embedded_Test_bold.otf', key: '', subsetted: false},
		{name: 'Zz Embedded Test', slot: 3, path: 'fonts/Zz_Embedded_Test_bolditalic.ttf', key: '', subsetted: false}
	]);
	assert.equal(result.fontAltNames.Arial, 'Arimo');
	assert.equal(result.fontFamilies['Zz Embedded Test'], 'swiss');
});

test('paths that are not files of media/fonts are ignored', () => {
	const bytes = fontTable([[strItem(1, 'X'), strItem(4, '../secret.ttf'), strItem(5, 'fonts/../a.ttf'),
		strItem(6, 'image1.png'), strItem(7, 'http://example.com/a.ttf'), strItem(4, 'fonts/ok.ttf')]]);
	const {result} = readTable(cur, bytes);
	eq(result.embeddedFonts, [{name: 'X', slot: 0, path: 'fonts/ok.ttf', key: '', subsetted: false}]);
});

test('a font table without embedded fonts is read exactly as before', () => {
	const old = loadReader(committed('word/Editor/Serialize2.js'));
	const bytes = fontTable([[strItem(1, 'Arial'), strItem(2, 'Arimo'), strItem(3, 'swiss')], [strItem(1, 'Cursive'), strItem(3, 'script')]]);
	const a = readTable(old, bytes), b = readTable(cur, bytes);
	assert.equal(a.res, b.res);
	eq(a.result.fontAltNames, b.result.fontAltNames);
	eq(a.result.fontFamilies, b.result.fontFamilies);
	assert.equal(b.result.embeddedFonts.length, 0);
});

test('an old reader skips the new items (they are unknown items of a font)', () => {
	const old = loadReader(committed('word/Editor/Serialize2.js'));
	const bytes = fontTable([[strItem(1, 'Zz'), strItem(2, 'Yy'), strItem(4, 'fonts/a.ttf'), strItem(5, 'fonts/b.ttf')]]);
	const {res, result} = readTable(old, bytes);
	assert.equal(res, AC.c_oSerConstants.ReadOk);
	assert.equal(result.fontAltNames.Zz, 'Yy');
});

// ---- registration in the font engine ----
const SOURCE_NAMES = ['SetDocumentEmbeddedFonts', 'HasPendingEmbeddedFonts', 'ClearDocumentEmbeddedFonts', 'RegisterEmbeddedFonts',
	'LoadPendingEmbeddedFonts', 'GetFontFile', 'GetFontFileWeb'];

function engine(options) {
	options = options || {};
	const A = {g_font_files: [], g_font_infos: [], g_map_font_index: {}};
	A.CFontFileLoader = function (id) {
		this.Id = id;
		this.Status = -1;
		this.LoadFontFromData = function (data) {
			this.data = data;
			this.Status = 0;
		};
	};
	A.CFontInfo = function (name, thumb, iR, fR, iI, fI, iB, fB, iBI, fBI) {
		Object.assign(this, {Name: name, indexR: iR, faceIndexR: fR, indexI: iI, faceIndexI: fI, indexB: iB, faceIndexB: fB, indexBI: iBI, faceIndexBI: fBI});
	};
	// an installed family "Installed Sans" with regular and italic only
	A.g_font_files.push({Id: 'inst-r'}, {Id: 'inst-i'});
	A.g_font_infos.push(new A.CFontInfo('Installed Sans', 0, 0, 0, 1, 0, -1, -1, -1, -1));
	A.g_map_font_index['Installed Sans'] = 0;
	const requests = [];
	function XHR() {
		this.open = (m, url) => {
			this.url = url;
		};
		this.send = () => {
			requests.push(this.url);
			const r = options.responses[this.url];
			setTimeout(() => {
				if (r === undefined) return this.onerror();
				this.status = 200;
				this.response = r.buffer.slice(r.byteOffset, r.byteOffset + r.byteLength);
				this.onload();
			}, 0);
		};
	}
	const app = {FontPickerMap: {x: 1}, DocumentEmbedded: {}, EmbeddedPending: [], EmbeddedUndo: [], GetFontFileOld: null};
	const ctx = vm.createContext({
		AscFonts: A, window: {}, console, XMLHttpRequest: XHR, Uint8Array, setTimeout, CFontSelectFormat: function () {},
		AscCommon: {g_oDocumentUrls: {getImageUrl: p => options.noUrls ? null : 'http://doc/media/' + p}}
	});
	for (const n of SOURCE_NAMES) {
		app[n] = vm.runInContext('(' + assignedFunction(mapSrc, n) + ')', ctx);
	}
	app.GetSubstituteName = n => n;
	app.GetFontIndex = () => ({m_wsFontName: 'default'});
	return {A, app, requests};
}
const sfnt = (tag, n) => {
	const d = new Uint8Array(n || 64);
	d.set(tag, 0);
	return d;
};
const TT = [0, 1, 0, 0];

test('a family that is not installed is registered with its faces and found by the font selection', () => {
	const {A, app} = engine();
	const touched = app.RegisterEmbeddedFonts([
		{name: 'Zz', slot: 0, data: sfnt(TT)}, {name: 'Zz', slot: 1, data: sfnt(TT)}, {name: 'Zz', slot: 2, data: sfnt(TT)}
	]);
	eq(touched, ['Zz']);
	const info = A.g_font_infos[A.g_map_font_index.Zz];
	assert.equal(info.Name, 'Zz');
	assert(info.isEmbeddedFont);
	assert(info.indexR >= 0 && info.indexB >= 0 && info.indexI >= 0);
	assert.equal(info.indexBI, -1);
	assert.equal(A.g_font_files[info.indexR].Status, 0);
	eq(app.DocumentEmbedded.Zz, {m_wsFontName: 'Zz'});
	eq(app.FontPickerMap, {});
	eq(app.GetFontFile('Zz', 0), {m_wsFontName: 'Zz'});
});

test('an installed family keeps its styles; the embedded face only fills a missing style', () => {
	const {A, app} = engine();
	const info = A.g_font_infos[0];
	app.RegisterEmbeddedFonts([
		{name: 'Installed Sans', slot: 0, data: sfnt(TT)}, {name: 'Installed Sans', slot: 1, data: sfnt(TT)},
		{name: 'Installed Sans', slot: 2, data: sfnt(TT)}
	]);
	assert.equal(info.indexR, 0);		// installed regular wins
	assert.equal(info.indexI, 1);		// installed italic wins
	assert(info.indexB >= 2);			// bold was missing: embedded
	assert.equal(info.indexBI, -1);
	assert.equal(A.g_font_infos.length, 1, 'no second family');
	eq(app.DocumentEmbedded, {});
});

test('subsetted font without regular face: the first face stands for regular', () => {
	const {A, app} = engine();
	app.RegisterEmbeddedFonts([{name: 'OnlyBold', slot: 1, data: sfnt(TT)}]);
	const info = A.g_font_infos[A.g_map_font_index.OnlyBold];
	assert.equal(info.indexR, info.indexB);
	assert.equal(info.indexI, -1);
});

test('clearing drops the registration of the document and restores the installed font', () => {
	const {A, app} = engine();
	const before = JSON.stringify(A.g_font_infos);
	app.RegisterEmbeddedFonts([{name: 'Zz', slot: 0, data: sfnt(TT)}, {name: 'Installed Sans', slot: 1, data: sfnt(TT)}]);
	app.ClearDocumentEmbeddedFonts();
	assert.equal(A.g_map_font_index.Zz, undefined);
	assert.equal(JSON.stringify(A.g_font_infos), before);
	eq(app.DocumentEmbedded, {});
	eq(app.GetFontFile('Zz', 0), {m_wsFontName: 'default'});
});

test('reading the next document drops the fonts of the previous one', () => {
	const {A, app} = engine();
	app.RegisterEmbeddedFonts([{name: 'Zz', slot: 0, data: sfnt(TT)}]);
	app.SetDocumentEmbeddedFonts([]);
	assert.equal(A.g_map_font_index.Zz, undefined);
	assert(!app.HasPendingEmbeddedFonts());
});

test('invalid faces (no data, bad slot, no name) register nothing', () => {
	const {A, app} = engine();
	eq(app.RegisterEmbeddedFonts([{name: 'Zz', slot: 0}, {name: 'Zz', slot: 9, data: sfnt(TT)}, {slot: 0, data: sfnt(TT)}, null]), []);
	assert.equal(A.g_font_infos.length, 1);
});

test('pending fonts are fetched from the document media, validated, then registered', async () => {
	const responses = {
		'http://doc/media/fonts/a.ttf': sfnt(TT),
		'http://doc/media/fonts/b.otf': sfnt([0x4F, 0x54, 0x54, 0x4F]),
		'http://doc/media/fonts/junk.ttf': Uint8Array.from({length: 100}, (_, i) => i + 7),
		'http://doc/media/fonts/short.ttf': sfnt(TT, 5)
	};
	const {A, app, requests} = engine({responses});
	app.SetDocumentEmbeddedFonts([
		{name: 'Zz', slot: 0, path: 'fonts/a.ttf'}, {name: 'Zz', slot: 1, path: 'fonts/b.otf'},
		{name: 'Bad', slot: 0, path: 'fonts/junk.ttf'}, {name: 'Bad2', slot: 0, path: 'fonts/short.ttf'},
		{name: 'Missing', slot: 0, path: 'fonts/none.ttf'}
	]);
	assert(app.HasPendingEmbeddedFonts());
	await new Promise(resolve => app.LoadPendingEmbeddedFonts(resolve));
	assert.equal(requests.length, 5);
	assert(A.g_map_font_index.Zz !== undefined);
	assert.equal(A.g_map_font_index.Bad, undefined);
	assert.equal(A.g_map_font_index.Bad2, undefined);
	assert.equal(A.g_map_font_index.Missing, undefined);
	assert(!app.HasPendingEmbeddedFonts());
});

test('when nothing can be loaded the callback still runs and nothing is registered (as without the feature)', async () => {
	const {A, app} = engine({noUrls: true, responses: {}});
	app.SetDocumentEmbeddedFonts([{name: 'Zz', slot: 0, path: 'fonts/a.ttf'}]);
	await new Promise(resolve => app.LoadPendingEmbeddedFonts(resolve));
	assert.equal(A.g_map_font_index.Zz, undefined);
	const failing = engine({responses: {}});
	failing.app.SetDocumentEmbeddedFonts([{name: 'Zz', slot: 0, path: 'fonts/a.ttf'}]);
	await new Promise(resolve => failing.app.LoadPendingEmbeddedFonts(resolve));
	assert.equal(failing.A.g_map_font_index.Zz, undefined);
});

test('LoadDocumentFonts waits for the embedded fonts, and is unchanged without them', () => {
	const src = assignedFunction(loaderSrc, 'LoadDocumentFonts');
	const run = pending => {
		const calls = [];
		let pendingCb = null;
		const ctx = vm.createContext({
			g_fontApplication: {
				HasPendingEmbeddedFonts: () => pending,
				LoadPendingEmbeddedFonts: cb => {
					pendingCb = cb;
				}
			}
		});
		const loader = {LoadDocumentFontsNow: fonts => calls.push(fonts)};
		loader.LoadDocumentFonts = vm.runInContext('(' + src + ')', ctx);
		loader.LoadDocumentFonts('F');
		return {calls, pendingCb};
	};
	const without = run(false);
	eq(without.calls, ['F']);
	const withFonts = run(true);
	eq(withFonts.calls, []);
	withFonts.pendingCb();
	eq(withFonts.calls, ['F']);
});

test('embedded families are not offered in the font list', () => {
	assert(/info\.Name !== "ASCW3" && !info\.isEmbeddedFont/.test(loaderSrc));
});

test('a font embedded in the document is loaded under the name of the document, not the family name of the file', () => {
	// text shaping looks a font up again by the family name of the loaded face: for an embedded font that name
	// must be the registered one (the file of a document may carry another family name than the font table says)
	const externals = read('common/Drawings/Externals.js');
	assert(/\(this\.isEmbeddedFont \|\| \(AscFonts\.DesktopFontPolicy && this\.isLocalFont\)\) \? \{name: this\.Name, style: style\} : null/.test(externals));
});
