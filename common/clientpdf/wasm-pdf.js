/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function (window) {
	'use strict';
	var BASE = new URL('../../../../sdkjs/common/clientpdf/', window.location.href).href;
	function unsupported(code, detail) { return AscCommon.ClientPdf.unsupported(code, detail); }
	function join(parts) {
		var out = new Uint8Array(parts.reduce(function (n, p) { return n + p.length; }, 0)), at = 0;
		parts.forEach(function (p) { out.set(p, at); at += p.length; }); return out;
	}
	// Scan resource references without interpreting or rasterizing any drawing command.
	function Scanner() {
		this.fonts = new Map(); this.images = new Map(); this.font = 'Arial'; this.style = 0;
		this.pages = 0; this.links = 0;
	}
	Scanner.prototype.feed = function (bytes) {
		var u = bytes.slice(), dv = new DataView(u.buffer), p = 0, last = 0, parts = [], self = this;
		function skip(n) { if (!Number.isSafeInteger(n) || n < 0 || p + n > u.length) throw unsupported('command', 'truncated stream'); p += n; }
		function byte() { skip(1); return u[p - 1]; }
		function int() { skip(4); return dv.getInt32(p - 4, true); }
		function str(wide) {
			var n; if (wide) n = int(); else { skip(2); n = dv.getUint16(p - 2, true) * 2; }
			if (n % 2) throw unsupported('command', 'invalid UTF-16');
			var at = p; skip(n); var s = '';
			for (var i = at; i < p; i += 2) s += String.fromCharCode(dv.getUint16(i, true));
			return s;
		}
		function font() { self.fonts.set(self.font + '|' + self.style, { name: self.font, style: self.style }); }
		function image(wide) {
			var start = p, path = str(wide), entry = self.images.get(path);
			if (!entry) {
				var ext = (/\.([a-z0-9]+)(?:[?#]|$)/i.exec(path) || [null, 'png'])[1].toLowerCase();
				entry = { source: path, path: 'media/image' + self.images.size + '.' + ext }; self.images.set(path, entry);
			}
			var name = entry.path, b = new Uint8Array((wide ? 4 : 2) + name.length * 2), v = new DataView(b.buffer);
			if (wide) v.setInt32(0, name.length * 2, true); else v.setUint16(0, name.length, true);
			for (var j = 0; j < name.length; j++) v.setUint16((wide ? 4 : 2) + j * 2, name.charCodeAt(j), true);
			parts.push(u.subarray(last, start), b); last = p;
		}
		while (p < u.length) {
			var c = byte();
			switch (c) {
				case 202: this.pages++; this.font = 'Arial'; this.style = 0; break;
				case 203: case 33: case 97: case 98: case 100: break;
				case 1: case 3: case 21: case 22: case 23: case 42: case 99: case 121: case 122: case 200: case 201: skip(4); break;
				case 2: case 5: case 6: case 7: case 24: case 25: case 27: case 28: case 30: skip(1); break;
				case 4: if (byte() === 5) skip(int() * 4); break;
				case 26: image(false); break;
				case 32: image(true); break;
				case 29: skip(16); break;
				case 31:
					byte();
					for (;;) {
						var sub = byte(); if (sub === 251) break;
						if (sub === 0) skip(21); else if (sub === 1) skip(25);
						else if (sub === 2) skip(int() * 8); else throw unsupported('gradient', String(sub));
					} break;
				case 41: this.font = str(false); break;
				case 43: this.style = int(); break;
				case 80: str(false); skip(8); font(); break;
				case 83: skip(12); skip(int() * 4); font(); break;
				case 91: case 92: case 104: case 105: skip(8); break;
				case 94: case 130: skip(24); break;
				case 111: image(true); skip(16); break;
				case 160: skip(16); str(false); str(false); this.links++; break;
				case 161: skip(28); this.links++; break;
				case 163:
					var flags = int(); for (var k = 0; k < 4; k++) if (flags & (1 << k)) str(false); break;
				case 162: throw unsupported('form-fields');
				default:
					if (c < 160) throw unsupported('command', String(c));
					var length = int(); if (length < 4) throw unsupported('command', 'invalid length'); skip(length - 4); break;
			}
		}
		parts.push(u.subarray(last)); return join(parts);
	};
	function embeddingAllowed(bytes, face) {
		try {
			var fsType = window.EOWasmFonts.inspectFace(bytes, face).embedding;
			if (fsType & (2 | 0x200 | 0x100)) throw unsupported('font-license');
		} catch (e) { if (e.eoUnsupported) throw e; throw unsupported('font', 'invalid font data'); }
	}

	function cancelled() { return Object.assign(new Error('PDF export cancelled'), {name:'AbortError',code:'aborted'}); }
	function check(signal) { if(signal && signal.aborted) throw cancelled(); }
	function progress(env,value) { try { if(env.onProgress) { var result=env.onProgress(value); if(result&&result.catch)result.catch(function(){}); } } catch(_error) {} }
	function abortable(promise,signal) {
		if(!signal)return promise;
		return new Promise(function(resolve,reject){
			var done=false;
			function finish(error,value){if(done)return;done=true;signal.removeEventListener('abort',cancel);error?reject(error):resolve(value);}
			function cancel(){finish(cancelled());}
			signal.addEventListener('abort',cancel,{once:true});
			Promise.resolve(promise).then(function(value){finish(null,value);},function(error){finish(error);});
			if(signal.aborted)cancel();
		});
	}
	function convert(job, runtime, env) {
		return new Promise(function (resolve, reject) {
			var worker,timer,done=false,signal=env.signal;
			function finish(error,result){if(done)return;done=true;clearTimeout(timer);if(signal)signal.removeEventListener('abort',cancel);if(worker)worker.terminate();error?reject(error):resolve(result);}
			function cancel(){finish(cancelled());}
			if(signal&&signal.aborted){cancel();return;}
			try { worker = new Worker(BASE + 'wasm-worker.js'); } catch (e) { finish(unsupported('wasm-worker', e.message)); return; }
			timer=setTimeout(function(){finish(unsupported('wasm-timeout'));},120000);
			worker.onerror=function(event){if(event&&event.preventDefault)event.preventDefault();finish(unsupported('wasm-worker','worker failed'));};
			worker.onmessage=function(event){
				if(done)return;var m=event.data;
				if(m&&m.type==='progress'){if(['loading-runtime','initializing','converting'].indexOf(m.stage)>=0)progress(env,{stage:m.stage});return;}
				if(!m||!(m.bytes instanceof Uint8Array)&&!m.error){finish(unsupported('wasm-worker','invalid worker response'));return;}
				if(m.error)finish(unsupported(m.code||'wasm',m.error));else finish(null,m);
			};
			if(signal)signal.addEventListener('abort',cancel,{once:true});
			job.runtime=runtime;job.progress=typeof env.onProgress==='function';
			var transfers=[job.stream.buffer];job.files.forEach(function(f){transfers.push(f.bytes.buffer);});
			try{worker.postMessage(job,transfers);}catch(e){finish(unsupported('wasm-worker',e.message));}
		});
	}
	function Writer(env) { this.env = env; this.scanner = new Scanner(); this.parts = []; this.stats = {}; }
	Writer.prototype.feed = async function (u) { check(this.env.signal);this.parts.push(this.scanner.feed(u)); };
	var fontLibrary;
	function loadFonts() {
		if (!fontLibrary) fontLibrary = new Promise(function (resolve, reject) {
			if (window.EOWasmFonts) { resolve(window.EOWasmFonts); return; }
			var script = document.createElement('script');
			script.src = BASE + 'wasm-fonts.js';
			script.onload = function () { window.EOWasmFonts ? resolve(window.EOWasmFonts) : reject(unsupported('font', 'shared font service missing')); };
			script.onerror = function () { reject(unsupported('font', 'shared font service unavailable')); };
			document.head.appendChild(script);
		});
		return fontLibrary;
	}
	Writer.prototype.finish = async function () {
		try{
		var files = [], fonts = [], fontBytes = 0, imageBytes = 0;
		check(this.env.signal);progress(this.env,{stage:'preparing'});
		var service = await abortable(loadFonts(),this.env.signal);check(this.env.signal);
		// The same prepareFS interface is used by local conversion modules. Collect
		// isolated copies here; the PDF worker receives them through transferables.
		var selected = await abortable(service.prepareFS({ FS: {
			mkdirTree: function () {},
			writeFile: function (path, bytes) { files.push({ path: path, bytes: bytes.slice() }); }
		} }, Array.from(this.scanner.fonts.values())),this.env.signal);check(this.env.signal);
		fontBytes = selected.bytes;
		selected.written.forEach(function (font) {
			embeddingAllowed(font.bytes, font.faceIndex || 0);
			fonts.push({ name: font.name, style: font.style, source: font.source });
		});
		for (var image of this.scanner.images.values()) {
			check(this.env.signal);
			var raw;
			try { raw = await abortable(this.env.getRawImage(image.source),this.env.signal); } catch (e) { if(e.name==='AbortError')throw e;throw unsupported('image', image.source); }
			if (!raw || !raw.bytes.length) throw unsupported('image', image.source);
			imageBytes += raw.bytes.length; files.push({ path: '/p/' + image.path, bytes: raw.bytes.slice() });
			// The native renderer uses the original vector sidecar when one exists.
			if (/\.svg$/i.test(image.path)) {
				for (var ext of ['emf', 'wmf']) {
					var side = await abortable(this.env.getRawImage(image.source.replace(/\.svg(?=[?#]|$)/i, '.' + ext)),this.env.signal).catch(function (error) { if(error.name==='AbortError')throw error;return null; });
					if (side) { imageBytes += side.bytes.length; files.push({ path: '/p/' + image.path.replace(/\.svg$/i, '.' + ext), bytes: side.bytes.slice() }); }
				}
			}
		}
		check(this.env.signal);
		var result;
		try{result=await convert({stream:join(this.parts),files:files},this.env['wasmRuntime'],this.env);}finally{this.parts=[];}
		check(this.env.signal);
		this.parts = [];
		this.stats = { pages: this.scanner.pages, images: this.scanner.images.size, links: this.scanner.links,
			fonts: fonts, fontBytes: fontBytes, fontFiles: new Set(selected.written.map(function (font) { return font.path; })).size,
			fontDownloadedBytes: selected.written.reduce(function (n, font) { return n + (font.downloadBytes || 0); }, 0),
			imageBytes: imageBytes, worker: true,
			convertMs: result.convertMs, initMs: result.initMs, heapBytes: result.heapBytes };
		return result.bytes;
		}finally{this.parts=[];}
	};
	window.EOWasmPdf = window['EOWasmPdf'] = { Writer: Writer, Scanner: Scanner };
})(window);
