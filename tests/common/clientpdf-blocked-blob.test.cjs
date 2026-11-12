// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const source = fs.readFileSync(path.join(root, 'common/clientpdf/clientpdf.js'), 'utf8');

class Memory {
	constructor() { this.bytes = []; }
	Seek() { this.bytes.length = 0; }
	GetCurPosition() { return this.bytes.length; }
	GetDataUint8() { return Uint8Array.from(this.bytes); }
	ClearNoAttack() { this.bytes.length = 0; }
	write(value) { this.bytes.push(...Buffer.from(value)); }
}

class Renderer {
	constructor() { this.Memory = new Memory(); }
	InitPicker() {}
	DocInfo() { this.Memory.write('META'); }
	BeginPage() {}
	EndPage() {}
	AddHeadings() { this.Memory.write('HEAD'); }
}

function loadClientPdf(downloadImpl, frame) {
	const downloadCalls = [];
	const AscCommon = {
		CDocumentRenderer: Renderer,
		CMemory: Memory,
		g_oTextMeasurer: {m_oManager: {}},
		g_oDocumentUrls: {},
		changeFileExtention: (name, ext) => String(name).replace(/\.[^.]+$/, '') + '.' + ext,
		DownloadType: {None: 0, Print: 1},
		DownloadFileFromBytes: function (bytes, filename, mime, opt_bom, opt_onBlocked) {
			downloadCalls.push({filename, mime, hasCallback: typeof opt_onBlocked === 'function'});
			if (downloadImpl) downloadImpl(opt_onBlocked);
		}
	};
	const window = {
		AscCommon,
		AscWord: {},
		localStorage: {getItem: () => null, setItem() {}},
		location: {href: 'https://local.invalid/editor/'},
		performance: {now: () => Date.now()}
	};
	if (frame) { window.self = {}; window.top = {}; }
	const context = {
		window,
		AscCommon,
		AscFonts: {g_fonts_streams: [], g_map_font_index: {}, g_font_infos: []},
		Asc: {
			c_oAscDisplayModeInReview: {Edit: 0, Final: 1, Original: 2, Simple: 3},
			c_oAscFileType: {PDF: 2},
			c_oAscAsyncAction: {DownloadAs: 1, Print: 2},
			c_oAscAsyncActionType: {BlockInteraction: 0}
		},
		Uint8Array, ArrayBuffer, Buffer, Promise, Date, URL, Blob,
		fetch: async () => ({ok: true, arrayBuffer: async () => new ArrayBuffer(0)}),
		createImageBitmap: async () => ({}),
		performance: window.performance,
		setTimeout, clearTimeout, setInterval, clearInterval, console
	};
	if (frame) context.navigator = {userAgent: frame};
	vm.runInNewContext(source, context);
	const client = window.AscCommon.ClientPdf;
	client.registerBackend({
		id: 'test',
		load: async () => {},
		create: () => ({
			stats: {pages: 1},
			nextPageIndex: 0,
			feed: async function (u8) { this.stats.bytes = (this.stats.bytes || 0) + u8.length; },
			finish: async () => new Uint8Array([37, 80, 68, 70])
		})
	});
	client.setBackend('test');
	return {client, downloadCalls};
}

function makeApi() {
	const downloadAsCalls = [];
	const logic = {
		GetDisplayModeInReview: () => 0,
		SetDisplayModeInReview() {},
		DrawPage(page, renderer) { renderer.Memory.write('page'); }
	};
	const drawing = {
		m_lPagesCount: 1,
		m_lCountCalculatePages: 1,
		m_arrPages: [{width_mm: 210, height_mm: 297}],
		m_oLogicDocument: logic
	};
	return {
		WordControl: {m_oDrawingDocument: drawing, m_oLogicDocument: logic},
		documentTitle: 'doc.docx',
		ShowParaMarks: true,
		asc_getCoreProps: () => ({}),
		asc_isFinalizedVersion: () => false,
		sendEvent() {},
		sync_EndAction() {},
		downloadAs: function (actionType, options) { downloadAsCalls.push({actionType, options}); },
		downloadAsCalls
	};
}

test('blocked blob download: server re-export once with eoClientPdfSkip, next export skips client build', async () =>
{
	// simulate the page CSP blocking the blob: link: the callback fires
	const {client, downloadCalls} = loadClientPdf((cb) => { if (cb) cb(); });
	const api = makeApi();
	const options = {fileType: 2};

	assert.equal(client.canHandle(api, 1, options, 0), true);
	client.run(api, 1, options, 0);
	await new Promise(resolve => setImmediate(resolve));
	await new Promise(resolve => setImmediate(resolve));

	// the client-built bytes were handed to the download once, with a callback
	assert.equal(downloadCalls.length, 1);
	assert.equal(downloadCalls[0].hasCallback, true);
	assert.equal(downloadCalls[0].filename, 'doc.pdf');
	// the callback re-ran the export through the server exactly once, with the skip flag
	assert.equal(api.downloadAsCalls.length, 1);
	assert.equal(api.downloadAsCalls[0].actionType, 1);
	assert.equal(api.downloadAsCalls[0].options.eoClientPdfSkip, true);
	// the next export goes straight to the server: canHandle is false
	assert.equal(client.canHandle(api, 1, {fileType: 2}, 0), false);
});

test('working blob download: no server re-export, flag stays clear', async () =>
{
	const {client, downloadCalls} = loadClientPdf(null);   // callback never fires
	const api = makeApi();
	const options = {fileType: 2};

	client.run(api, 1, options, 0);
	await new Promise(resolve => setImmediate(resolve));
	await new Promise(resolve => setImmediate(resolve));

	assert.equal(downloadCalls.length, 1);
	assert.equal(api.downloadAsCalls.length, 0);
	assert.equal(client.canHandle(api, 1, {fileType: 2}, 0), true);
});

test('embedded in a frame: Firefox and Safari go to the server, Chromium keeps the client build', () =>
{
	const api = makeApi();
	const firefox = loadClientPdf(null, 'Mozilla/5.0 (X11; Linux x86_64; rv:130.0) Gecko/20100101 Firefox/130.0').client;
	assert.equal(firefox.canHandle(api, 1, {fileType: 2}, 0), false);
	const safari = loadClientPdf(null, 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15').client;
	assert.equal(safari.canHandle(api, 1, {fileType: 2}, 0), false);
	const chrome = loadClientPdf(null, 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36').client;
	assert.equal(chrome.canHandle(api, 1, {fileType: 2}, 0), true);
	const edge = loadClientPdf(null, 'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0').client;
	assert.equal(edge.canHandle(api, 1, {fileType: 2}, 0), true);
});
