/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
(function (window) {
	'use strict';
	function unsupported(code, detail) { return failure(code + ': ' + (detail || '')); }
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
	var CACHE = 'eo-wasm-fonts-v1', pending = new Map();
	function failure(detail) {
		var e = new Error('client PDF: unsupported font (' + detail + ')');
		e.eoUnsupported = true; e.code = 'font'; e.detail = detail; return e;
	}
	function requests(names) {
		var unique = new Map();
		Array.from(names || []).forEach(function (entry) {
			var name = typeof entry === 'string' ? entry : entry.name;
			var styles = typeof entry === 'string' ? [0, 1, 2, 3] : [entry.style || 0];
			if (typeof name !== 'string' || !name.trim()) throw failure('invalid font name');
			styles.forEach(function (style) {
				if (!Number.isInteger(style) || style < 0 || style > 3) throw failure('invalid font style');
				unique.set(name + '|' + style, { name: name, style: style });
			});
		});
		return Array.from(unique.values());
	}
	async function consent(lf) {
		if (!lf || !lf.hasConsent || !lf.hasConsent()) return false;
		var bridge;
		try { bridge = window.eoDesktop || window.parent && window.parent.eoDesktop; } catch (e) {}
		if (bridge && typeof bridge.fontConsent === 'function') {
			try { return typeof bridge.fontPermission === 'function' && await bridge.fontPermission(); }
			catch (e) { return false; }
		}
		if (window.navigator && window.navigator.permissions) {
			try { return (await window.navigator.permissions.query({ name: 'local-fonts' })).state === 'granted'; }
			catch (e) { return false; }
		}
		return false;
	}
	function streamBytes(index) {
		var st = window.AscFonts.g_fonts_streams[index];
		if (!st) return null;
		var n = st.size === undefined ? st.len : st.size, data = st.data;
		if (typeof data === 'number') return window.AscFonts.GetUint8ArrayFromPointer(data, n).slice();
		if (data instanceof ArrayBuffer) return new Uint8Array(data, 0, n);
		if (data && data.length) return data.subarray ? data.subarray(0, n) : Uint8Array.from(data).subarray(0, n);
		return null;
	}
	async function load(entry, options) {
		var A = window.AscFonts, common = window.AscCommon;
		if (!A || !common || !common.g_font_loader) throw failure('editor font index unavailable');
		var idx = A.g_map_font_index[entry.name], info = idx === undefined ? null : A.g_font_infos[idx];
		if (!info) throw failure(entry.name + ': not in editor font index');
		var id = info.GetFontID(common.g_font_loader, entry.style);
		if (!id || !id.file) throw failure(entry.name + ': no font face');
		var lf = window.Common && window.Common.ElixirLocalFonts;
		var local = !!(info.isLocalFont || id.file.localFontData || (lf && lf.isLocal && lf.isLocal(entry.name)));
		if (local && !(await consent(lf))) throw failure(entry.name + ': local font consent not granted');
		var epoch = lf && lf.desktopPolicy && typeof lf.desktopPolicy.epoch === 'function' ? lf.desktopPolicy.epoch() : undefined;
		var need = info.GetNeedInfo ? info.GetNeedInfo(entry.style) : {};
		var meta = { name: entry.name, style: entry.style, faceIndex: id.faceIndex || 0,
			key: id.file.Id + '#' + (id.faceIndex || 0), source: local ? 'local' : 'server',
			needBold: !!need.needB, needItalic: !!need.needI };
		var bytes = id.file.Status === 0 ? streamBytes(id.file.stream_index) : null;
		var cache, cacheKey;
		if (!local) {
			var url = new URL((common.g_font_loader.fontFilesPath || '') + id.file.Id, window.location.href).href;
			cacheKey = url + (url.indexOf('?') < 0 ? '?' : '&') + 'eo-wasm-decoded=1';
			try { cache = window.caches && await window.caches.open(CACHE); } catch (e) {}
			if (!bytes && cache) {
				var hit = await cache.match(cacheKey);
				if (hit) { bytes = new Uint8Array(await hit.arrayBuffer()); meta.cacheHit = true; }
			}
			if (!bytes) {
				var response = await fetch(url, { credentials: 'same-origin' });
				if (!response.ok) throw failure(entry.name + ': HTTP ' + response.status);
				bytes = new Uint8Array(await response.arrayBuffer());
				// FontLoader.LoadFontArrayBuffer decodes this same 32-byte ODTTF prefix.
				var guid = [160,102,214,32,20,150,71,250,149,105,184,80,176,65,73,72];
				for (var i = 0; i < Math.min(32, bytes.length); i++) bytes[i] ^= guid[i % 16];
				meta.downloadBytes = bytes.length;
			}
			if (cache && !meta.cacheHit) {
				try { await cache.put(cacheKey, new Response(bytes)); } catch (e) {}
			}
		} else {
			// Reuse the editor's consented lazy loader; never enumerate or prompt here.
			var source = options.getFont || common.ClientPdf && common.ClientPdf.fonts.get;
			var font = source && await source(entry.name, entry.style);
			if (font) {
				if (font.variableUnsupported) throw failure(entry.name + ': nondefault variable instance requires static embedding');
				bytes = font.bytes; meta.faceIndex = font.faceIndex || 0; meta.key = font.key;
				meta.fileKey = font.fileKey || id.file.localFontData && id.file.localFontData.fileId || font.key.split('#')[0];
				meta.needBold = font.needBold === undefined ? meta.needBold : !!font.needBold;
				meta.needItalic = font.needItalic === undefined ? meta.needItalic : !!font.needItalic;
			} else bytes = null;
			if (!(await consent(lf)) || epoch !== undefined && lf.desktopPolicy.epoch() !== epoch) throw failure(entry.name + ': local font consent changed');
		}
		if (!bytes || !bytes.length) throw failure(entry.name + ': font data unavailable');
		if (bytes.length > 128 * 1024 * 1024) throw failure(entry.name + ': font exceeds 128 MiB limit');
		meta.bytes = bytes; return meta;
	}
	async function ensureFonts(names, options) {
		options = options || {};
		var written = [], missing = [], files = new Map(), bytes = 0;
		// Sequential reads keep transient memory bounded by the selected set.
		for (var entry of requests(names)) {
			var key = entry.name + '|' + entry.style, promise = pending.get(key);
			try {
				if (!promise) {
					promise = load(entry, options); pending.set(key, promise);
					promise.finally(function (k) { return function () { pending.delete(k); }; }(key)).catch(function () {});
				}
				var font = await promise;
				var fileKey = font.fileKey || font.key.split('#')[0];
				if (!files.has(fileKey)) {
					if (bytes + font.bytes.length > 192 * 1024 * 1024) throw failure('selected fonts exceed 192 MiB limit');
					files.set(fileKey, '/fonts/font' + files.size + '.ttf');
					bytes += font.bytes.length;
				}
				written.push(Object.assign({}, font, { path: files.get(fileKey) }));
			} catch (e) { missing.push({ name: entry.name, style: entry.style, reason: e.detail || e.message }); }
		}
		return { written: written, missing: missing, bytes: bytes };
	}
	// FreeType encodes named instance + 1 in high16, collection face in low16.
	function inspectFace(data, faceIndex) {
		if (!Number.isInteger(faceIndex) || faceIndex < 0 || faceIndex > 0x7fffffff) throw failure('invalid encoded face index');
		var dv = new DataView(data.buffer, data.byteOffset, data.byteLength), collection = faceIndex & 65535, instance = faceIndex >>> 16;
		function bounds(at, n) { if (!Number.isSafeInteger(at) || at < 0 || n < 0 || at + n > data.length) throw failure('invalid font bounds'); }
		bounds(0, 12);
		var offset = 0;
		if (dv.getUint32(0) === 0x74746366) {
			var count = dv.getUint32(8);
			if (!count || count > 64 || collection >= count) throw failure('invalid collection face');
			bounds(12, count * 4); offset = dv.getUint32(12 + collection * 4);
		} else if (collection) throw failure('noncollection face index');
		bounds(offset, 12);
		var signature = dv.getUint32(offset), n = dv.getUint16(offset + 4), tables = {};
		if ([0x00010000, 0x4f54544f, 0x74727565].indexOf(signature) < 0 || n > 256) throw failure('invalid SFNT');
		bounds(offset + 12, n * 16);
		for (var i = 0; i < n; i++) {
			var at = offset + 12 + i * 16, start = dv.getUint32(at + 8), length = dv.getUint32(at + 12);
			bounds(start, length);
			tables[String.fromCharCode.apply(null, data.subarray(at, at + 4))] = {start: start, length: length};
		}
		function table(tag, min) { var t = tables[tag]; if (t && t.length < min) throw failure('truncated ' + tag); return t && t.start; }
		var head = table('head', 54), os = table('OS/2', 78), post = table('post', 16), fvar = table('fvar', 16);
		if (os !== undefined) table('OS/2', dv.getUint16(os) >= 2 ? 96 : dv.getUint16(os) >= 1 ? 86 : 78);
		var physical = head === undefined ? 0 : dv.getUint16(head + 44);
		var weight = os === undefined ? 400 : dv.getUint16(os + 4), axes = {}, nonDefault = false;
		if (fvar !== undefined) {
			var axisOffset = dv.getUint16(fvar + 4), axisCount = dv.getUint16(fvar + 8), axisSize = dv.getUint16(fvar + 10);
			var instanceCount = dv.getUint16(fvar + 12), instanceSize = dv.getUint16(fvar + 14), end = fvar + tables.fvar.length;
			if (axisOffset < 16 || axisCount > 64 || axisSize < 20 || instanceCount > 4096 || instanceSize < 4 + axisCount * 4 ||
				fvar + axisOffset + axisCount * axisSize + instanceCount * instanceSize > end || instance > instanceCount) throw failure('invalid variable face');
			for (var a = 0; a < axisCount; a++) {
				var axis = fvar + axisOffset + a * axisSize;
				var tag = String.fromCharCode.apply(null, data.subarray(axis, axis + 4));
				var value = instance ? fvar + axisOffset + axisCount * axisSize + (instance - 1) * instanceSize + 4 + a * 4 : axis + 8;
				axes[tag] = dv.getInt32(value) / 65536;
				if (dv.getInt32(value) !== dv.getInt32(axis + 8)) nonDefault = true;
				var min = dv.getInt32(axis + 4) / 65536, max = dv.getInt32(axis + 12) / 65536;
				if (axes[tag] < min || axes[tag] > max || min > max) throw failure('invalid variable coordinate');
			}
		} else if (instance) throw failure('named instance without fvar');
		if (axes.wght !== undefined && (axes.wght < 1 || axes.wght > 1000)) throw failure('invalid variable weight');
		if (axes.wght !== undefined) { weight = Math.round(axes.wght); physical = (physical & ~1) | (weight >= 700 ? 1 : 0); }
		if (axes.ital !== undefined || axes.slnt !== undefined) physical = (physical & ~2) | (axes.ital || axes.slnt ? 2 : 0);
		return {offset: offset, tables: tables, head: head, os: os, post: post, physical: physical, weight: weight,
			embedding: os === undefined ? 0 : dv.getUint16(os + 8), axes: axes, nonDefault: nonDefault};
	}
	function selection(fonts) {
		var parts = [], total = 4;
		function record(font) {
			var chunks = [], size = 0;
			function add(bytes) { chunks.push(bytes); size += bytes.length; }
			function num(value, n) {
				var u = new Uint8Array(n), v = new DataView(u.buffer);
				if (n === 2) v.setInt16(0, value, true); else v.setInt32(0, value, true); add(u);
			}
			function str(value) { var u = new TextEncoder().encode(value); num(u.length, 4); add(u); }
			var data = font.bytes, dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
			var inspected = inspectFace(data, font.faceIndex), offset = inspected.offset;
			var desktop = window.Common && window.Common.ElixirLocalFonts && window.Common.ElixirLocalFonts.desktopPolicy;
			if (desktop && font.source === 'local' && inspected.nonDefault) throw failure(font.name + ': nondefault variable instance requires static embedding');
			var os = inspected.os, head = inspected.head, post = inspected.post;
			function u16(at, fallback) { return at === undefined ? fallback : dv.getUint16(at); }
			function s16(at, fallback) { return at === undefined ? fallback : dv.getInt16(at); }
			function at(base, n) { return base === undefined ? undefined : base + n; }
			num(0, 4); str(font.name); num(0, 4); str(font.path);
			// Default named aliases are the same instance; native readers require the collection ordinal.
			num(desktop && font.source === 'local' && !inspected.nonDefault ? font.faceIndex & 65535 : font.faceIndex, 4);
			var style = inspected.physical;
			num(!!(style & 2), 4); num(!!(style & 1), 4);
			num(post === undefined ? 0 : dv.getUint32(post + 12) !== 0, 4);
			num(10, 4); add(os === undefined ? new Uint8Array(10) : data.slice(os + 32, os + 42));
			for (var j = 0; j < 6; j++) {
				var p = j < 4 ? at(os, 42 + j * 4) : os !== undefined && dv.getUint16(os) >= 1 ? os + 78 + (j - 4) * 4 : undefined;
				num(p === undefined ? 0 : dv.getUint32(p), 4);
			}
			num(inspected.weight, 2); num(u16(at(os, 6), 5), 2);
			num(s16(at(os, 30), 0), 2); num(dv.getUint32(offset) === 0x4f54544f ? 2 : 1, 2);
			var scale = 1000 / u16(at(head, 18), 1000);
			[2, 68, 70, 72].forEach(function (n) { num(Math.trunc(s16(at(os, n), 0) * scale), 2); });
			num(os !== undefined && dv.getUint16(os) >= 2 ? Math.trunc(dv.getInt16(os + 86) * scale) : 0, 2);
			num(os !== undefined && dv.getUint16(os) >= 2 ? Math.trunc(dv.getInt16(os + 88) * scale) : 0, 2);
			num(u16(at(os, 8), 0), 2);
			var out = new Uint8Array(size), pos = 0;
			chunks.forEach(function (c) { out.set(c, pos); pos += c.length; });
			new DataView(out.buffer).setInt32(0, size, true); return out;
		}
		var unique = new Map();
		fonts.forEach(function (f) { unique.set(f.name + '|' + f.path + '|' + f.faceIndex, f); });
		unique.forEach(function (f) { var u = record(f); parts.push(u); total += u.length; });
		var lf = window.Common && window.Common.ElixirLocalFonts, policy = lf && lf.desktopPolicy;
		var ranges = policy && policy.nativeRanges ? policy.nativeRanges(fonts) : [];
		if (ranges.length) {
			if (ranges.length > 100000) throw failure('too many selected font ranges');
			var rangeParts = [], length = 4;
			ranges.forEach(function (range) {
				var name = new TextEncoder().encode(range.name), row = new Uint8Array(12 + name.length), view = new DataView(row.buffer);
				view.setInt32(0, name.length, true); row.set(name, 4);
				view.setInt32(4 + name.length, range.start, true); view.setInt32(8 + name.length, range.end, true);
				rangeParts.push(row); length += row.length;
			});
			var suffix = new Uint8Array(length), at = 4; new DataView(suffix.buffer).setInt32(0, ranges.length, true);
			rangeParts.forEach(function (row) { suffix.set(row, at); at += row.length; });
			parts.push(suffix); total += suffix.length;
		}
		var result = new Uint8Array(total), pos = 4;
		new DataView(result.buffer).setInt32(0, unique.size, true);
		parts.forEach(function (u) { result.set(u, pos); pos += u.length; }); return result;
	}
	async function prepareFS(module, names) {
		var result = await ensureFonts(names), FS = module.FS;
		if (result.missing.length) throw failure(result.missing.map(function (x) { return x.reason; }).join('; '));
		FS.mkdirTree('/fonts');
		var paths = new Set();
		result.written.forEach(function (font) {
			if (!paths.has(font.path)) { FS.writeFile(font.path, font.bytes); paths.add(font.path); }
		});
		FS.writeFile('/font_selection.bin', selection(result.written));
		return result;
	}
	function fontsOfDocument(input) {
		if (Array.isArray(input)) {
			var unique = new Map();
			input.flatMap(fontsOfDocument).forEach(function (entry) {
				var key = typeof entry === 'string' ? 'family:' + entry : 'style:' + entry.name + '|' + entry.style;
				unique.set(key, entry);
			});
			return Array.from(unique.values());
		}
		var u = typeof input === 'string' ? new TextEncoder().encode(input) : input instanceof Uint8Array ? input : new Uint8Array(input);
		if (u[0] === 202) {
			var scanner = new Scanner(); scanner.feed(u);
			return Array.from(scanner.fonts.values());
		}
		// Word editor binaries have a typed font table. Do not guess by searching strings.
		var text = new TextDecoder().decode(u.subarray(0, 32)), header = /^DOCY;v(\d+);(\d+);/.exec(text);
		if (!header) throw failure('unsupported document font stream');
		var base = header[0].length, data = u, start = base;
		if (Number(header[1]) !== 10) {
			var encoded = typeof input === 'string' ? input.slice(base) : new TextDecoder().decode(u.subarray(base));
			var decoded = window.atob(encoded); data = Uint8Array.from(decoded, function (c) { return c.charCodeAt(0); }); start = 0;
		}
		var dv = new DataView(data.buffer, data.byteOffset, data.byteLength), names = new Set(), found = false;
		for (var i = 0; i < data[start]; i++) {
			var p = start + 1 + i * 5;
			if (data[p] !== 22) continue;
			found = true; p = dv.getUint32(p + 1, true);
			var end = p + 4 + dv.getUint32(p, true); p += 4;
			if (end > data.length) throw failure('truncated document font table');
			while (p < end) {
				var type = data[p++], length = dv.getUint32(p, true); p += 4;
				var next = p + length;
				if (next > end) throw failure('truncated document font record');
				if (type === 0) while (p < next) {
					var field = data[p++], n = dv.getUint32(p, true); p += 4;
					if (p + n > next || n % 2) throw failure('invalid document font name');
					if (field === 1) names.add(new TextDecoder('utf-16le').decode(data.subarray(p, p + n)));
					p += n;
				}
				p = next;
			}
		}
		if (!found) throw failure('document font table unavailable; use drawing stream');
		return Array.from(names);
	}
	window.EOWasmFonts = window['EOWasmFonts'] = {
		ensureFonts: ensureFonts, prepareFS: prepareFS, fontsOfDocument: fontsOfDocument,
		cacheName: CACHE, selection: selection, inspectFace: inspectFace
	};
})(window);
