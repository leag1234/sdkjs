/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Client-side PDF export (Word editor). Download as PDF and Print produce the PDF in the browser
 * from the editor's own drawing-command stream (the same AscCommon.CDocumentRenderer stream the
 * server would otherwise replay), with the font files the screen uses (AscFonts.g_fonts_streams,
 * including fonts read from the user's computer). Nothing is sent to the server.
 *
 * The native wasm PDF worker is loaded on first use only; the JS writer and fontkit
 * are loaded only if the wasm backend fails. Whatever the writer cannot reproduce makes the export fall back to the server
 * path, with an event that lets the interface say why (asc_onClientPdf {state:'fallback'}).
 *
 * Two interfaces keep the native wasm and in-house JS PDF writers replaceable
 * (see docs/client-pdf.md):
 *  - ClientPdf.fonts (font source): get(name, style) -> Promise<{key, bytes, faceIndex, needBold,
 *    needItalic, source: 'server'|'local'}>: the font file bytes the screen uses, whether they came from
 *    the server or from the user's computer (Local Font Access);
 *  - PdfBackend: {id, load() -> Promise, create(env) -> {feed(u8) -> Promise, finish() -> Promise<Uint8Array>,
 *    stats}}; env = {fonts, getFont, getImage, title, lang, producer, creator}. The backend is fed the
 *    drawing-command stream page by page; an error with eoUnsupported = true and a code makes the
 *    export fall back to the server. ClientPdf.registerBackend(backend), ClientPdf.setBackend(id).
 *
 * Events: api.sendEvent('asc_onClientPdf', {state: 'start'|'progress'|'done'|'fallback'|'error', ...}).
 * Switch: AscCommon.ClientPdf.setEnabled(false) (interface option) sends every PDF to the server.
 */
(function (window) {
	'use strict';

	var BASE = '../../../../sdkjs/common/clientpdf/';
	var MAX_PAGES = 3000;          // above this, the server path (memory of the tab)
	var YIELD_MS = 40;             // give the browser a frame at least every YIELD_MS of work

	// Scripts covered by the bundled pan-CJK faces. Both PDF backends use this one predicate.
	var CJK_RANGES = [[0x1100, 0x11FF], [0x2E80, 0xA4CF], [0xAC00, 0xD7FF],
		[0xF900, 0xFAFF], [0xFF61, 0xFF9F], [0x20000, 0x323AF]];
	function isCjkCodePoint(codepoint) {
		for (var i = 0; i < CJK_RANGES.length; i++)
			if (codepoint >= CJK_RANGES[i][0] && codepoint <= CJK_RANGES[i][1]) return true;
		return false;
	}

	var enabled = true;
	try { if (window.localStorage && window.localStorage.getItem('elixir.clientpdf') === 'off') enabled = false; } catch (e) { /* private mode */ }

	var libs = null;
	function loadScript(url) {
		return new Promise(function (resolve, reject) {
			AscCommon.loadScript(url, resolve, function () { reject(new Error('cannot load ' + url)); });
		});
	}
	function loadLibs() {
		if (libs) return libs;
		libs = (async function () {
			if (!window.fontkit) {
				// the vendored fontkit is a UMD build: hide the AMD loader of the interface while it registers
				var def = window.define, amd = def && def.amd;
				if (amd) def.amd = undefined;
				try { await loadScript(BASE + 'vendor/fontkit/fontkit.umd.min.js'); }
				finally { if (amd) def.amd = amd; }
			}
			if (!window.EOPdf) await loadScript(BASE + 'eo-pdf.js');
			if (!window.fontkit || !window.EOPdf) throw new Error('client PDF libraries missing');
		})();
		libs.catch(function () { libs = null; });
		return libs;
	}

	function now() { return window.performance ? performance.now() : Date.now(); }
	function nextFrame() { return new Promise(function (r) { setTimeout(r, 0); }); }

	// ---- fonts: the bytes the screen already uses ------------------------------------------------
	function streamBytes(index) {
		var st = AscFonts.g_fonts_streams[index];
		if (!st) return null;
		var len = st.size !== undefined ? st.size : st.len;
		var d = st.data;
		if (typeof d === 'number') return AscFonts.GetUint8ArrayFromPointer(d, len).slice();
		if (d instanceof Uint8Array) return d.subarray(0, len);
		if (d instanceof ArrayBuffer) return new Uint8Array(d, 0, len);
		if (d && d.length) return Uint8Array.from(Array.prototype.slice.call(d, 0, len));
		return null;
	}
	function ensureLoaded(file) {
		if (file.Status === 0 && file.stream_index >= 0) return Promise.resolve(true);
		return new Promise(function (resolve) {
			var done = false, finish = function () { if (!done) { done = true; resolve(file.Status === 0); } };
			var base = (AscCommon.g_font_loader && AscCommon.g_font_loader.fontFilesPath) || '';
			try {
				if (file.LoadFontAsync(base, finish) === true && file.Status === 0) finish();
			} catch (e) { finish(); }
			var t = setInterval(function () { if (file.Status === 0 || file.Status === 1) { clearInterval(t); finish(); } }, 50);
			setTimeout(function () { clearInterval(t); finish(); }, 20000);
		});
	}
	async function getFont(name, style) {
		var idx = AscFonts.g_map_font_index[name];
		if (idx === undefined) return null;
		var info = AscFonts.g_font_infos[idx], loader = AscCommon.g_font_loader;
		var lf = window.Common && window.Common.ElixirLocalFonts;
		var id = info.GetFontID(loader, style), epoch = lf && lf.desktopPolicy && typeof lf.desktopPolicy.epoch === 'function' && lf.desktopPolicy.epoch();
		if (!id || !id.file) return null;
		var local = !!(info.isLocalFont || id.file.localFontData || lf && lf.isLocal && lf.isLocal(name));
		async function permitted() {
			if (!local) return true;
			if (!lf || !lf.hasConsent || !lf.hasConsent()) return false;
			var bridge = lf.desktopBridge && lf.desktopBridge();
			if (bridge && (!bridge.fontPermission || !(await bridge.fontPermission()))) return false;
			return !lf.desktopPolicy || typeof lf.desktopPolicy.epoch !== 'function' || lf.desktopPolicy.epoch() === epoch;
		}
		if (!(await permitted())) return null;
		// A lazy collection/variable loader updates the selected face after reading bytes.
		for (var attempt = 0; attempt < 3; attempt++) {
			if (!(await ensureLoaded(id.file)) || !(await permitted())) return null;
			var current = info.GetFontID(loader, style);
			if (!current || !current.file) return null;
			if (current.file !== id.file) { id = current; continue; }
			id = current;
			var bytes = streamBytes(id.file.stream_index);
			if (!bytes) return null;
			var need = info.GetNeedInfo ? info.GetNeedInfo(style) : {};
			var fileKey = id.file.localFontData && id.file.localFontData.fileId || id.file.Id;
			return { key: fileKey + '#' + (id.faceIndex || 0), fileKey: fileKey,
				bytes: bytes, faceIndex: id.faceIndex || 0,
				variableUnsupported: !!(local && lf && lf.desktopPolicy && id.file.localFontData &&
					id.file.localFontData.variable && (id.faceIndex >>> 16) && id.file.localFontData.instanceDefault !== true),
				needBold: !!need.needB, needItalic: !!need.needI,
				source: local ? 'local' : 'server' };
		}
		return null;
	}
	var FontSource = { get: getFont };

	function unsupported(code, detail) {
		var e = new Error('client PDF: unsupported ' + code + (detail ? ' (' + detail + ')' : ''));
		e.eoUnsupported = true; e.code = code; e.detail = detail || '';
		return e;
	}

	// ---- backends --------------------------------------------------------------------------------
	var backends = {}, backendId = 'js', preferredBackend = 'wasm';
	var wasmLib = null, runtimeProbe = null, runtimeUrl = null;
	// One discovery attempt per editor session, including absent or invalid runtimes.
	function discoverRuntime() {
		if (!runtimeProbe) runtimeProbe = (async function () {
			try {
				var response = await fetch(BASE + 'runtime/current.json', { cache: 'no-store' });
				if (!response.ok) return false;
				var pointer = await response.json(), version = pointer.artifact_version;
				if (typeof version !== 'string' || !/^[0-9a-f]+-[0-9a-f]+-st$/.test(version)) return false;
				var url = BASE + 'runtime/' + version + '/';
				response = await fetch(url + 'manifest.json');
				if (!response.ok) return false;
				var manifest = await response.json();
				if (manifest.artifact_version !== version || !manifest.build_flags || manifest.build_flags.module_kind !== 'pdf') return false;
				runtimeUrl = new URL(url, window.location.href).href;
				registerBackend({
					id: 'wasm',
					load: function () {
						if (!wasmLib) {
							wasmLib = loadScript(BASE + 'wasm-pdf.js').then(function () {
								if (!window.EOWasmPdf) throw unsupported('wasm', 'backend missing');
							});
							wasmLib.catch(function () { wasmLib = null; });
						}
						return wasmLib;
					},
					create: function (env) { env['wasmRuntime'] = runtimeUrl; return new window.EOWasmPdf.Writer(env); }
				});
				return true;
			} catch (e) { return false; }
		})();
		return runtimeProbe;
	}
	function registerBackend(b) { backends[b.id] = b; }
	registerBackend({
		id: 'js',   // in-house writer on fontkit (common/clientpdf/eo-pdf.js)
		load: loadLibs,
		create: function (env) { env.fontkit = window.fontkit; return new window.EOPdf.Writer(env); }
	});

	// ---- images: the files the screen already loaded (browser cache), never uploaded -------------
	function jpegInfo(b) {
		var i = 2, adobe = false;
		while (i + 9 < b.length) {
			if (b[i] !== 0xFF) { i++; continue; }
			var m = b[i + 1], len = (b[i + 2] << 8) | b[i + 3];
			if (m === 0xEE) adobe = true;
			if ((m >= 0xC0 && m <= 0xC3) || (m >= 0xC5 && m <= 0xC7) || (m >= 0xC9 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF))
				return { height: (b[i + 5] << 8) | b[i + 6], width: (b[i + 7] << 8) | b[i + 8], comps: b[i + 9], adobe: adobe };
			if (m === 0xD8 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
			i += 2 + len;
		}
		return null;
	}
	function imageUrl(path) {
		if (/^(data:|blob:|https?:)/i.test(path)) return path;
		var u = AscCommon.g_oDocumentUrls;
		return (u.getImageUrl && u.getImageUrl(path)) || (u.getUrl && u.getUrl(path)) || path;
	}
	function loadElement(url) {
		return new Promise(function (resolve, reject) {
			var img = new Image();
			img.crossOrigin = 'anonymous';
			img.onload = function () { resolve(img); };
			img.onerror = function () { reject(new Error('image')); };
			img.src = url;
		});
	}
	async function getImage(path, wMm, hMm) {
		var url = imageUrl(path);
		var resp = await fetch(url, { credentials: 'same-origin', cache: 'force-cache' });
		if (!resp.ok) return null;
		var blob = await resp.blob();
		var bytes = new Uint8Array(await blob.arrayBuffer());
		if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
			var ji = jpegInfo(bytes);
			if (ji && (ji.comps === 1 || ji.comps === 3 || ji.comps === 4)) return { kind: 'jpeg', bytes: bytes, width: ji.width, height: ji.height, comps: ji.comps, adobe: ji.adobe };
		}
		var isSvg = /svg/i.test(blob.type) || (bytes[0] === 0x3C /* < */);
		var src, w, h;
		if (isSvg) {
			var burl = URL.createObjectURL(new Blob([bytes], { type: 'image/svg+xml' }));
			try { src = await loadElement(burl); } finally { URL.revokeObjectURL(burl); }
			// vector images are rasterized at 300 dpi of their size on the page (capped)
			var scale = 300 / 25.4;
			w = Math.max(1, Math.min(4000, Math.round((wMm || 50) * scale)));
			h = Math.max(1, Math.min(4000, Math.round((hMm || 50) * scale)));
		} else {
			src = await createImageBitmap(blob, { imageOrientation: 'none', premultiplyAlpha: 'none', colorSpaceConversion: 'none' });
			w = src.width; h = src.height;
		}
		var cv = document.createElement('canvas');
		cv.width = w; cv.height = h;
		var ctx = cv.getContext('2d');
		ctx.drawImage(src, 0, 0, w, h);
		if (src.close) src.close();
		var px = ctx.getImageData(0, 0, w, h).data, n = w * h;
		var rgb = new Uint8Array(n * 3), alpha = new Uint8Array(n), opaque = true;
		for (var i = 0, j = 0; i < n; i++, j += 4) {
			rgb[i * 3] = px[j]; rgb[i * 3 + 1] = px[j + 1]; rgb[i * 3 + 2] = px[j + 2];
			alpha[i] = px[j + 3];
			if (px[j + 3] !== 255) opaque = false;
		}
		return { kind: 'raw', width: w, height: h, rgb: rgb, alpha: opaque ? null : alpha };
	}

	async function getRawImage(path) {
		var response = await fetch(imageUrl(path), { credentials: 'same-origin', cache: 'force-cache' });
		if (!response.ok) return null;
		return { bytes: new Uint8Array(await response.arrayBuffer()) };
	}

	function checkCancelled(signal){if(signal&&signal.aborted)throw Object.assign(new Error('PDF export cancelled'),{name:'AbortError',code:'aborted'});}
	function reportProgress(options,value){try{if(options.onProgress){var result=options.onProgress(value);if(result&&result.catch)result.catch(function(){});}}catch(_error){}}

	// ---- capture + write -------------------------------------------------------------------------
	async function buildOnce(api, options, onProgress, id) {
		checkCancelled(options.signal);
		var backend = backends[id];
		if (!backend) throw unsupported('backend', id);
		await backend.load();checkCancelled(options.signal);
		var t0 = now();
		var dd = api.WordControl.m_oDrawingDocument;
		var logic = api.WordControl.m_oLogicDocument;
		var mergedPrint = false;
		if (options.oDocumentMailMerge && (!dd.GenerateMailMergePrint || !dd.GenerateMailMergePrint(options.oDocumentMailMerge)))
			throw unsupported('mail-merge-layout');
		mergedPrint = !!options.oDocumentMailMerge;
		try {
		var adv = options.advancedOptions;
		var isSelection = !!(adv && adv.asc_getPrintType && Asc.c_oAscPrintType && Asc.c_oAscPrintType.Selection === adv.asc_getPrintType());
		if (isSelection) dd.GenerateSelectionPrint();
		var src = dd.printedDocument ? dd.printedDocument.DrawingDocument : dd;
		var count = Math.min(src.m_lPagesCount, src.m_lCountCalculatePages);
		var native = adv && adv.asc_getNativeOptions ? adv.asc_getNativeOptions() : null;
		var ranges = native && native['pages'] !== undefined && !isSelection ? AscCommon.getNativePrintRanges(native['pages'], native['currentPage'], count) : undefined;
		var list = [];
		for (var p = 0; p < count; p++) if (!ranges || ranges[p]) list.push(p);
		if (!list.length) throw unsupported('empty');
		if (list.length > MAX_PAGES) throw unsupported('size', list.length + ' pages');

		var writer = backend.create({
			fonts: FontSource, getFont: FontSource.get, getImage: getImage, getRawImage: getRawImage,
			title: api.documentTitle ? String(api.documentTitle).replace(/\.[^.]+$/, '') : '',
			lang: api.asc_getLocale ? (function () { try { var l = api.asc_getLocale(); return typeof l === 'string' ? l : ''; } catch (e) { return ''; } })() : '',
			producer: 'ElixirOffice', creator: 'ElixirOffice',signal:options.signal,onProgress:options.onProgress
		});

		var watermark = api.watermarkDraw;
		var r = new AscCommon.CDocumentRenderer();
		r.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
		r.VectorMemoryForPrint = new AscCommon.CMemory();
		r.isPrintMode = !!(options.isPdfPrint || (api.asc_isFinalizedVersion && api.asc_isFinalizedVersion()));
		var oldMarks = api.ShowParaMarks;
		var captureMs = 0, last = now();
		try {
			if (watermark) watermark.StartRenderer();
			r.Memory.Seek(0);
			r.DocInfo(api.asc_getCoreProps());
			await writer.feed(r.Memory.GetDataUint8(0, r.Memory.GetCurPosition()));
			reportProgress(options,{stage:'converting',completed:0,total:list.length});
			for (var k = 0; k < list.length; k++) {
				checkCancelled(options.signal);
				var i = list[k], page = src.m_arrPages[i], tc = now();
				r.Memory.Seek(0);
				r.VectorMemoryForPrint.ClearNoAttack();
				api.ShowParaMarks = false;   // only while drawing: the screen keeps its marks between pages
				try {
					r.BeginPage(page.width_mm, page.height_mm);
					src.m_oLogicDocument.DrawPage(i, r);
					if (watermark) watermark.DrawOnRenderer(r, page.width_mm, page.height_mm);
					r.EndPage();
				} finally { api.ShowParaMarks = oldMarks; }
				captureMs += now() - tc;
				writer.nextPageIndex = i;
				await writer.feed(r.Memory.GetDataUint8(0, r.Memory.GetCurPosition()));
				checkCancelled(options.signal);
				if (onProgress) onProgress(k + 1, list.length);
				reportProgress(options,{stage:'converting',completed:k+1,total:list.length});
				if (now() - last > YIELD_MS) { await nextFrame(); last = now(); }
			}
			r.Memory.Seek(0);
			r.AddHeadings(src.m_oLogicDocument || logic, true);
			if (r.Memory.GetCurPosition()) await writer.feed(r.Memory.GetDataUint8(0, r.Memory.GetCurPosition()));
		} finally {
			if (watermark) watermark.EndRenderer();
			api.ShowParaMarks = oldMarks;
			dd.printedDocument = null;
		}
		checkCancelled(options.signal);var bytes = await writer.finish();checkCancelled(options.signal);
		var stats = writer.stats || {};
		stats.backend = backend.id;
		stats.captureMs = Math.round(captureMs);
		stats.totalMs = Math.round(now() - t0);
		stats.bytes = bytes.length;
		return { bytes: bytes, stats: stats };
		} finally {
			if (mergedPrint)
				dd.printedDocument = null;
		}
	}

	async function build(api, options, onProgress) {
		if (window.AscWord && window.AscWord.CellWidowReplay)
			window.AscWord.CellWidowReplay.assertLayoutReady(api.WordControl && api.WordControl.m_oLogicDocument);
		var id = preferredBackend, start = now();
		if (id === 'wasm' && !(await discoverRuntime())) id = 'js';
		backendId = id;
		try { return await buildOnce(api, options, onProgress, id); }
		catch (error) {
			if (id !== 'wasm'||error.name==='AbortError'||error.code==='aborted'||options.signal&&options.signal.aborted) throw error;
			var result = await buildOnce(api, options, onProgress, 'js');
			result.stats.fallbacks = [{ from: 'wasm', to: 'js', reason: error.code || 'wasm', detail: error.detail || error.message }];
			result.stats.totalMs = Math.round(now() - start);
			return result;
		}
	}

	// ---- integration with downloadAs -------------------------------------------------------------
	var ClientPdf = {
		last: null,
		isEnabled: function () { return enabled; },
		setEnabled: function (v) {
			enabled = !!v;
			try { window.localStorage && window.localStorage.setItem('elixir.clientpdf', enabled ? 'on' : 'off'); } catch (e) { /* private mode */ }
		},
		// Word only for now; Download (to the host integration) and mail merge keep the server path
		canHandle: function (api, actionType, options, downloadType) {
			if (!enabled || options.eoClientPdfSkip) return false;
			if (Asc.c_oAscFileType.PDF !== options.fileType) return false;    // PDF/A: server
			if (window['AscDesktopEditor'] || window.isCloudCryptoDownloadAs) return false;
			if (options.oDocumentMailMerge || options.oMailMergeSendData || options.isSaveAs || options.wopiSaveAsPath) return false;
			if (actionType !== Asc.c_oAscAsyncAction.DownloadAs && actionType !== Asc.c_oAscAsyncAction.Print) return false;
			if (downloadType !== AscCommon.DownloadType.None && downloadType !== AscCommon.DownloadType.Print) return false;
			if (!api.WordControl || !api.WordControl.m_oLogicDocument || (api.isDocumentRenderer && api.isDocumentRenderer())) return false;
			if (api.DocInfo && api.DocInfo.isFormatWithForms && api.DocInfo.isFormatWithForms()) return false;
			if (typeof Promise === 'undefined' || typeof Uint8Array === 'undefined' || typeof fetch === 'undefined' || typeof createImageBitmap === 'undefined') return false;
			return true;
		},
		// returns true: the export is handled here (asynchronously); on an unsupported case the
		// server path is started again with the same options
		run: function (api, actionType, options, downloadType) {
			var t = this, title = AscCommon.changeFileExtention(api.documentTitle, 'pdf', Asc.c_nMaxDownloadTitleLen);
			api.sendEvent('asc_onClientPdf', { state: 'start' });
			build(api, options, function (done, total) {
				api.sendEvent('asc_onClientPdf', { state: 'progress', page: done, pages: total });
			}).then(function (res) {
				t.last = res.stats;
				if (actionType) api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, actionType);
				if (downloadType === AscCommon.DownloadType.Print) {
					var url = URL.createObjectURL(new Blob([res.bytes], { type: 'application/pdf' }));
					api.sendEvent(downloadType, url, Asc.c_oAscFileType.PDF);
					setTimeout(function () { URL.revokeObjectURL(url); }, 10 * 60 * 1000);
				} else {
					AscCommon.DownloadFileFromBytes(res.bytes, title, 'application/pdf');
				}
				api.sendEvent('asc_onClientPdf', { state: 'done', pages: res.stats.pages, bytes: res.stats.bytes, ms: res.stats.totalMs });
			}, function (err) {
				t.last = { error: String(err && err.message || err), code: err && err.code };
				if (actionType) api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, actionType);
				var replay = window.AscWord && window.AscWord.CellWidowReplay;
				if (replay && replay.reportLayoutFailure(api.WordControl && api.WordControl.m_oLogicDocument, err)) {
					api.sendEvent('asc_onClientPdf', { state: 'error', reason: 'layout', detail: err.message });
					return;
				}
				if (window.console) console.warn('[client PDF] falling back to the server:', err);
				api.sendEvent('asc_onClientPdf', { state: 'fallback', reason: (err && err.eoUnsupported) ? err.code : 'error', detail: err && (err.detail || err.message) || '' });
				options.eoClientPdfSkip = true;    // stays set: a delayed retry (layout not finished) must not loop
				api.downloadAs(actionType, options);
			});
			return true;
		},
		// for tests and tools: the PDF bytes of the current document, no download
		build: function (api, options) { return build(api || window.editor, options || { fileType: Asc.c_oAscFileType.PDF }); },
		fonts: FontSource,
		getImage: getImage, getRawImage: getRawImage,
		unsupported: unsupported,
		cjkRanges: CJK_RANGES,
		isCjkCodePoint: isCjkCodePoint,
		registerBackend: registerBackend,
		setBackend: function (id) { preferredBackend = id; backendId = id === 'wasm' && !runtimeUrl ? 'js' : id; },
		discoverRuntime: discoverRuntime,
		getBackend: function () { return backendId; },
		loadLibs: loadLibs
	};

	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].ClientPdf = ClientPdf;
	window['AscCommon']['ClientPdf'] = ClientPdf;
	ClientPdf['setEnabled'] = ClientPdf.setEnabled;
	ClientPdf['isEnabled'] = ClientPdf.isEnabled;
	ClientPdf['build'] = ClientPdf.build;
	ClientPdf['fonts'] = FontSource;
	FontSource['get'] = FontSource.get;
	ClientPdf['getImage'] = getImage;
	ClientPdf['unsupported'] = unsupported;
	ClientPdf['cjkRanges'] = CJK_RANGES;
	ClientPdf['isCjkCodePoint'] = isCjkCodePoint;
	ClientPdf['registerBackend'] = ClientPdf.registerBackend;
	ClientPdf['setBackend'] = ClientPdf.setBackend;
	ClientPdf['getBackend'] = ClientPdf.getBackend;
	ClientPdf['discoverRuntime'] = ClientPdf.discoverRuntime;
})(window);
