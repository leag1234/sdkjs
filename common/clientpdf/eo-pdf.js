/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Browser-side PDF writer fed by the sdkjs drawing-command stream (the bytes that
 * AscCommon.CDocumentRenderer writes, the same stream the server-side x2t/PdfWriter
 * replays in core/DesktopEditor/graphics/MetafileToRenderer.cpp).
 *
 *   var w = new EOPdf.Writer(env);
 *   await w.feed(bytes);      // any number of times; each chunk ends on a command boundary
 *   var pdf = await w.finish(); // Uint8Array, PDF 1.7
 *
 * env = {
 *   fontkit,                                      // vendored fontkit (MIT)
 *   getFont(name, style)  -> Promise<{key, bytes, faceIndex, needBold, needItalic} | null>,
 *   getImage(path, wMm, hMm) -> Promise<{kind:'jpeg', bytes, width, height, comps, adobe}
 *                                       | {kind:'raw', width, height, rgb, alpha} | null>,
 *   producer, title, warn(code, detail)
 * }
 *
 * Text is written by glyph identifier (the stream carries the glyph ids computed by the
 * editor's shaper), fonts are embedded as subsets (TrueType -> CIDFontType2, CFF ->
 * CIDFontType0C), every font gets a ToUnicode CMap built from the code points the stream
 * attaches to each glyph. Anything the writer cannot reproduce faithfully throws an
 * EOPdf.Unsupported error so the caller can fall back to the server instead of producing a
 * broken file. No network, no server.
 */
(function (root) {
	'use strict';

	var MM = 72 / 25.4;

	function Unsupported(code, detail) {
		var e = new Error('client PDF: unsupported ' + code + (detail ? ' (' + detail + ')' : ''));
		e.eoUnsupported = true;
		e.code = code;
		e.detail = detail || '';
		return e;
	}

	// ---- small helpers ---------------------------------------------------------------------------
	function num(v) {
		if (!isFinite(v)) v = 0;
		var r = Math.round(v * 1000) / 1000;
		if (r === 0) return '0';
		var s = String(r);
		if (s.indexOf('e') >= 0) s = r.toFixed(3);
		return s;
	}
	function num5(v) { // text matrix: keep more precision for scaling terms
		var r = Math.round(v * 100000) / 100000;
		return r === 0 ? '0' : String(r).indexOf('e') >= 0 ? r.toFixed(5) : String(r);
	}
	function rgbStr(v) {
		return num((v & 255) / 255) + ' ' + num(((v >> 8) & 255) / 255) + ' ' + num(((v >> 16) & 255) / 255);
	}
	function hex4(v) { var s = v.toString(16); return '0000'.slice(s.length) + s; }
	function utf16hex(str) {
		var h = 'FEFF';
		for (var i = 0; i < str.length; i++) h += hex4(str.charCodeAt(i));
		return '<' + h + '>';
	}
	function textString(str) {
		var ascii = true;
		for (var i = 0; i < str.length; i++) { var c = str.charCodeAt(i); if (c < 32 || c > 126) { ascii = false; break; } }
		if (ascii) return '(' + str.replace(/[\\()]/g, '\\$&') + ')';
		return utf16hex(str);
	}
	function cpsToUtf16Hex(cps) {
		var h = '';
		for (var i = 0; i < cps.length; i++) {
			var c = cps[i];
			if (c > 0xFFFF) { c -= 0x10000; h += hex4(0xD800 + (c >> 10)) + hex4(0xDC00 + (c & 0x3FF)); }
			else h += hex4(c);
		}
		return h;
	}
	function latin1(str) {
		var u = new Uint8Array(str.length);
		for (var i = 0; i < str.length; i++) u[i] = str.charCodeAt(i) & 255;
		return u;
	}
	function concat(parts, total) {
		if (total === undefined) { total = 0; for (var i = 0; i < parts.length; i++) total += parts[i].length; }
		var out = new Uint8Array(total), p = 0;
		for (var j = 0; j < parts.length; j++) { out.set(parts[j], p); p += parts[j].length; }
		return out;
	}
	function pdfName(s) {
		return String(s).replace(/[^A-Za-z0-9\-_.]/g, '');
	}
	function pdfDate(d) {
		function p(n) { return (n < 10 ? '0' : '') + n; }
		var off = -d.getTimezoneOffset(), sign = off >= 0 ? '+' : '-';
		off = Math.abs(off);
		return 'D:' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()) +
			sign + p(Math.floor(off / 60)) + "'" + p(off % 60) + "'";
	}

	async function deflate(bytes) {
		if (typeof CompressionStream === 'undefined') return null;
		var cs = new CompressionStream('deflate');
		var stream = new Blob([bytes]).stream().pipeThrough(cs);
		return new Uint8Array(await new Response(stream).arrayBuffer());
	}

	// ---- stream reader ---------------------------------------------------------------------------
	function Reader(u8) {
		this.u = u8; this.p = 0;
		this.dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
	}
	Reader.prototype.byte = function () { return this.u[this.p++]; };
	Reader.prototype.ushort = function () { var v = this.dv.getUint16(this.p, true); this.p += 2; return v; };
	Reader.prototype.int = function () { var v = this.dv.getInt32(this.p, true); this.p += 4; return v; };
	Reader.prototype.uint = function () { var v = this.dv.getUint32(this.p, true); this.p += 4; return v; };
	Reader.prototype.dbl = function () { return this.int() / 100000; };
	Reader.prototype.str16 = function (bytes) {
		var s = '', end = this.p + bytes;
		for (var i = this.p; i + 1 < end; i += 2) s += String.fromCharCode(this.dv.getUint16(i, true));
		this.p = end;
		return s;
	};
	Reader.prototype.str = function () { return this.str16(2 * this.ushort()); }; // WriteString
	Reader.prototype.str2 = function () { return this.str16(this.int()); };        // WriteString2
	Reader.prototype.more = function () { return this.p < this.u.length; };

	// ---- fonts -----------------------------------------------------------------------------------
	function FontEntry(writer, key, fk, label) {
		this.w = writer;
		this.key = key;
		this.fk = fk;
		this.label = label;
		this.subset = fk.createSubset();
		this.cff = !!fk['CFF '] || !!fk.cff;
		this.upem = fk.unitsPerEm || 1000;
		this.cids = new Map();    // original gid -> cid (= glyph index in the subset)
		this.uni = new Map();     // cid -> code points (first seen)
		this.widths = new Map();  // cid -> width in 1/1000 em
		this.adv = new Map();     // gid -> advance in em
		this.res = 'F' + (writer.fonts.size + 1);
		this.ref = writer.alloc();
		var tag = '', n = writer.fonts.size + 1;
		for (var i = 0; i < 6; i++) { tag = String.fromCharCode(65 + (n % 26)) + tag; n = Math.floor(n / 26) + i * 7; }
		this.tag = tag;
		this.cid(0, null);
	}
	FontEntry.prototype.advance = function (gid) {
		var a = this.adv.get(gid);
		if (a === undefined) {
			var g = null;
			try { g = this.fk.getGlyph(gid); } catch (e) { g = null; }
			a = g ? (g.advanceWidth || 0) / this.upem : 0;
			this.adv.set(gid, a);
		}
		return a;
	};
	FontEntry.prototype.cid = function (gid, cps) {
		var cid = this.cids.get(gid);
		if (cid === undefined) {
			if (gid >= (this.fk.numGlyphs || 65535)) gid = 0;
			cid = this.subset.includeGlyph(gid);
			this.cids.set(gid, cid);
			this.widths.set(cid, Math.round(this.advance(gid) * 1000));
		}
		if (cps && cps.length && !this.uni.has(cid)) this.uni.set(cid, cps);
		return cid;
	};
	FontEntry.prototype.gidForCodePoint = function (cp) {
		try { var g = this.fk.glyphForCodePoint(cp); return g ? g.id : 0; } catch (e) { return 0; }
	};
	FontEntry.prototype.serialize = function () {
		var sub = this.subset;
		return new Promise(function (resolve, reject) {
			var parts = [];
			try {
				var s = sub.encodeStream();
				s.on('data', function (b) { parts.push(b); });
				s.on('end', function () { resolve(concat(parts)); });
				s.on('error', reject);
			} catch (e) { reject(e); }
		});
	};
	FontEntry.prototype.write = async function () {
		var w = this.w, fk = this.fk, k = 1000 / this.upem;
		var bytes;
		try { bytes = await this.serialize(); }
		catch (e) { throw Unsupported('font-subset', this.label + ': ' + (e && e.message)); }
		// fontkit writes the CFF header's offSize as 11 (out of the 1..4 range): readers that check it
		// (poppler) reject the font. The field is not used to read the table: set a valid value.
		if (this.cff && bytes.length > 4 && bytes[0] === 1 && (bytes[3] < 1 || bytes[3] > 4)) { bytes = bytes.slice(); bytes[3] = 4; }
		var psName = pdfName(fk.postscriptName || fk.fullName || this.label) || 'Font';
		var baseFont = this.tag + '+' + psName;
		var bbox = fk.bbox || { minX: 0, minY: -200, maxX: 1000, maxY: 800 };
		var os2 = null; try { os2 = fk['OS/2']; } catch (e) { os2 = null; }
		var ascent = Math.round((fk.ascent || 800) * k), descent = Math.round((fk.descent || -200) * k);
		var capHeight = Math.round(((os2 && os2.capHeight) || fk.capHeight || fk.ascent || 700) * k);
		var italic = 0; try { italic = fk.italicAngle || 0; } catch (e) { italic = 0; }
		var flags = 4 | (italic ? 64 : 0);

		var fileRef = w.alloc();
		var compressed = await deflate(bytes);
		var fdict = '/Length ' + (compressed ? compressed.length : bytes.length) + (compressed ? ' /Filter /FlateDecode' : '') +
			(this.cff ? ' /Subtype /CIDFontType0C' : ' /Length1 ' + bytes.length);
		w.putStream(fileRef, fdict, compressed || bytes);

		var descRef = w.alloc();
		w.put(descRef, '<< /Type /FontDescriptor /FontName /' + baseFont + ' /Flags ' + flags +
			' /FontBBox [' + [bbox.minX * k, bbox.minY * k, bbox.maxX * k, bbox.maxY * k].map(Math.round).join(' ') + ']' +
			' /ItalicAngle ' + num(italic) + ' /Ascent ' + ascent + ' /Descent ' + descent + ' /CapHeight ' + capHeight +
			' /StemV 80 ' + (this.cff ? '/FontFile3 ' : '/FontFile2 ') + fileRef + ' 0 R >>');

		// widths: [cid [w ...] ...] for consecutive cids
		var cids = Array.from(this.widths.keys()).sort(function (a, b) { return a - b; });
		var W = [], i = 0;
		while (i < cids.length) {
			var start = cids[i], ws = [this.widths.get(start)];
			while (i + 1 < cids.length && cids[i + 1] === cids[i] + 1) { i++; ws.push(this.widths.get(cids[i])); }
			W.push(start + ' [' + ws.join(' ') + ']');
			i++;
		}
		var cidRef = w.alloc();
		w.put(cidRef, '<< /Type /Font /Subtype /' + (this.cff ? 'CIDFontType0' : 'CIDFontType2') + ' /BaseFont /' + baseFont +
			' /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ' + descRef + ' 0 R' +
			' /DW 1000 /W [' + W.join(' ') + ']' + (this.cff ? '' : ' /CIDToGIDMap /Identity') + ' >>');

		// ToUnicode
		var lines = [], ent = Array.from(this.uni.entries()).sort(function (a, b) { return a[0] - b[0]; });
		for (var j = 0; j < ent.length; j += 100) {
			var chunk = ent.slice(j, j + 100);
			lines.push(chunk.length + ' beginbfchar');
			for (var q = 0; q < chunk.length; q++) lines.push('<' + hex4(chunk[q][0]) + '> <' + cpsToUtf16Hex(chunk[q][1]) + '>');
			lines.push('endbfchar');
		}
		var cmap = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
			'/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
			'1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' + lines.join('\n') + '\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n';
		var tuRef = w.alloc();
		await w.putStreamDeflate(tuRef, '', latin1(cmap));

		w.put(this.ref, '<< /Type /Font /Subtype /Type0 /BaseFont /' + baseFont + ' /Encoding /Identity-H /DescendantFonts [' +
			cidRef + ' 0 R] /ToUnicode ' + tuRef + ' 0 R >>');
		return { name: baseFont, cff: this.cff, glyphs: this.cids.size };
	};

	// ---- writer ----------------------------------------------------------------------------------
	function Writer(env) {
		this.env = env || {};
		this.objs = [null];          // id -> Uint8Array (serialized body)
		this.pages = [];             // {ref, contentRef, w, h, res:{F,X,G,S}, annots:[]}
		this.pageRefs = [];          // pre-allocated page refs (links can point forward)
		this.fonts = new Map();      // key -> FontEntry
		this.fontByName = new Map(); // name|style -> {entry, needBold, needItalic}
		this.images = new Map();     // path -> {res, ref}
		this.gstates = new Map();    // key -> {res, ref}
		this.shadings = 0;
		this.info = {};
		this.headings = null;
		this.stats = { pages: 0, glyphs: 0, texts: 0, paths: 0, images: 0, links: 0, gradients: 0, skipped: {}, ms: 0 };
		this.st = this.freshState();
		this.page = null;
	}
	Writer.prototype.freshState = function () {
		return {
			pen: 0, penA: 255, pw: 0.3, dash: null, join: 0,
			brush: 0, brushA: 255, bType: 1000, tex: null, texA: 255, texMode: 0, rect: null, rectOn: false, grad: null,
			font: '', size: 10, style: 0, T: [1, 0, 0, 1, 0, 0]
		};
	};
	Writer.prototype.alloc = function () { this.objs.push(null); return this.objs.length - 1; };
	Writer.prototype.put = function (id, body) { this.objs[id] = latin1(id + ' 0 obj\n' + body + '\nendobj\n'); };
	Writer.prototype.putStream = function (id, dict, bytes) {
		this.objs[id] = concat([latin1(id + ' 0 obj\n<< ' + dict + ' >>\nstream\n'), bytes, latin1('\nendstream\nendobj\n')]);
	};
	Writer.prototype.putStreamDeflate = async function (id, dict, bytes) {
		var c = await deflate(bytes);
		if (c) this.putStream(id, dict + ' /Length ' + c.length + ' /Filter /FlateDecode', c);
		else this.putStream(id, dict + ' /Length ' + bytes.length, bytes);
	};
	Writer.prototype.pageRef = function (i) {
		while (this.pageRefs.length <= i) this.pageRefs.push(this.alloc());
		return this.pageRefs[i];
	};
	Writer.prototype.skip = function (what) { this.stats.skipped[what] = (this.stats.skipped[what] || 0) + 1; };

	Writer.prototype.gs = function (fillA, strokeA) {
		var key = num(fillA) + '/' + num(strokeA), g = this.gstates.get(key);
		if (!g) {
			g = { res: 'G' + (this.gstates.size + 1), ref: this.alloc() };
			this.put(g.ref, '<< /Type /ExtGState /ca ' + num(fillA) + ' /CA ' + num(strokeA) + ' >>');
			this.gstates.set(key, g);
		}
		this.page.res.G[g.res] = g.ref;
		return '/' + g.res + ' gs';
	};

	Writer.prototype.font = async function (name, style) {
		var k = name + '|' + style, f = this.fontByName.get(k);
		if (f !== undefined) return f;
		var r = this.env.getFont ? await this.env.getFont(name, style) : null;
		if (!r || !r.bytes) throw Unsupported('font-missing', name);
		if (r.variableUnsupported) throw Unsupported('font-variable-instance', name + ': nondefault variable instance requires static embedding');
		var entry = this.fonts.get(r.key);
		if (!entry) {
			var fk;
			try {
				fk = this.env.fontkit.create(r.bytes);
				if (fk && fk.fonts && fk.fonts.length) fk = fk.fonts[(r.faceIndex || 0) & 0xFFFF] || fk.fonts[0]; // collection (.ttc)
				var inst = ((r.faceIndex || 0) >> 16);
				if (inst > 0 && fk.namedVariations) {                // variable font: named instance (index + 1)
					var names = Object.keys(fk.namedVariations);
					if (names[inst - 1]) fk = fk.getVariation(fk.namedVariations[names[inst - 1]]);
				}
			} catch (e) {
				throw Unsupported('font-format', name + ': ' + (e && e.message));
			}
			if (!fk || typeof fk.createSubset !== 'function') throw Unsupported('font-format', name);
			var fsType = 0; try { fsType = (fk['OS/2'] && fk['OS/2'].fsType) || 0; } catch (e) { fsType = 0; }
			// fsType: 0x0002 restricted licence embedding (no embedding allowed), 0x0200 bitmap embedding only
			if ((fsType & 0x000F) === 0x0002 || (fsType & 0x0200)) throw Unsupported('font-restricted', name);
			if (fk.directory && fk.directory.tables && fk.directory.tables.CFF2) throw Unsupported('font-format', name + ': CFF2');
			entry = new FontEntry(this, r.key, fk, name);
			this.fonts.set(r.key, entry);
		}
		f = { entry: entry, bold: !!r.needBold, italic: !!r.needItalic };
		this.fontByName.set(k, f);
		return f;
	};

	Writer.prototype.image = async function (path, wMm, hMm) {
		var im = this.images.get(path);
		if (im !== undefined) return im;
		var r = null;
		try { r = this.env.getImage ? await this.env.getImage(path, wMm, hMm) : null; }
		catch (e) { throw Unsupported('image', String(path).slice(0, 80) + ': ' + (e && e.message)); }
		if (!r) throw Unsupported('image', String(path).slice(0, 80));
		var ref = this.alloc(), res = 'Im' + (this.images.size + 1);
		if (r.kind === 'jpeg') {
			var cs = r.comps === 1 ? '/DeviceGray' : r.comps === 4 ? '/DeviceCMYK' : '/DeviceRGB';
			this.putStream(ref, '/Type /XObject /Subtype /Image /Width ' + r.width + ' /Height ' + r.height + ' /ColorSpace ' + cs +
				' /BitsPerComponent 8 /Filter /DCTDecode' + (r.comps === 4 && r.adobe ? ' /Decode [1 0 1 0 1 0 1 0]' : '') +
				' /Length ' + r.bytes.length, r.bytes);
		} else {
			var smask = '';
			if (r.alpha) {
				var mref = this.alloc();
				await this.putStreamDeflate(mref, '/Type /XObject /Subtype /Image /Width ' + r.width + ' /Height ' + r.height +
					' /ColorSpace /DeviceGray /BitsPerComponent 8', r.alpha);
				smask = ' /SMask ' + mref + ' 0 R';
			}
			await this.putStreamDeflate(ref, '/Type /XObject /Subtype /Image /Width ' + r.width + ' /Height ' + r.height +
				' /ColorSpace /DeviceRGB /BitsPerComponent 8' + smask, r.rgb);
		}
		im = { res: res, ref: ref };
		this.images.set(path, im);
		this.stats.images++;
		return im;
	};

	Writer.prototype.shading = function (g, T) {
		// g: {type:'lin'|'rad', p:[...], stops:[[pos, color]]}; coordinates transformed to page space (mm)
		var stops = g.stops.slice().sort(function (a, b) { return a[0] - b[0]; });
		if (!stops.length) return null;
		if (stops.length === 1) stops.push([1, stops[0][1]]);
		var c = function (v) { return rgbStr(v & 0xFFFFFF); };
		var fns = [], bounds = [], enc = [];
		for (var i = 0; i + 1 < stops.length; i++) {
			fns.push('<< /FunctionType 2 /Domain [0 1] /C0 [' + c(stops[i][1]) + '] /C1 [' + c(stops[i + 1][1]) + '] /N 1 >>');
			if (i > 0) bounds.push(num(stops[i][0]));
			enc.push('0 1');
		}
		var d0 = stops[0][0], d1 = stops[stops.length - 1][0];
		if (d1 <= d0) d1 = d0 + 1e-3;
		var fn = fns.length === 1 ? fns[0] : '<< /FunctionType 3 /Domain [' + num(d0) + ' ' + num(d1) + '] /Functions [' + fns.join(' ') +
			'] /Bounds [' + bounds.join(' ') + '] /Encode [' + enc.join(' ') + '] >>';
		var p = g.p, tp = function (x, y) { return [T[0] * x + T[2] * y + T[4], T[1] * x + T[3] * y + T[5]]; };
		var coords;
		if (g.type === 'lin') {
			var a = tp(p[0], p[1]), b = tp(p[2], p[3]);
			coords = '/ShadingType 2 /Coords [' + [a[0], a[1], b[0], b[1]].map(num).join(' ') + ']';
		} else {
			var s = Math.sqrt(Math.abs(T[0] * T[3] - T[1] * T[2])) || 1, a2 = tp(p[0], p[1]), b2 = tp(p[2], p[3]);
			coords = '/ShadingType 3 /Coords [' + [a2[0], a2[1], p[4] * s, b2[0], b2[1], p[5] * s].map(num).join(' ') + ']';
		}
		var ref = this.alloc(), res = 'Sh' + (++this.shadings);
		this.put(ref, '<< ' + coords + ' /ColorSpace /DeviceRGB /Function ' + fn + ' /Domain [' + num(d0) + ' ' + num(d1) + '] /Extend [true true] >>');
		this.page.res.S[res] = ref;
		this.stats.gradients++;
		return res;
	};

	// ---- page content ----------------------------------------------------------------------------
	function Page(writer, index) {
		this.w = 210; this.h = 297;
		this.index = index;
		this.out = [];
		this.res = { F: {}, X: {}, G: {}, S: {} };
		this.annots = [];
		this.path = [];
		this.clipping = false;
		this.top = { fill: '', gs: '', tr: 0, lw: '' };
		this.run = null;   // open text run
	}

	Writer.prototype.flushRun = function () {
		var pg = this.page, r = pg.run;
		if (!r) return;
		pg.run = null;
		var o = pg.out, top = pg.top;
		if (top.fill !== r.fill) { o.push(r.fill + ' rg'); top.fill = r.fill; }
		if (top.gs !== r.gs) { o.push(r.gs || this.gs(1, 1)); top.gs = r.gs; }
		if (r.bold) {
			if (top.stroke !== r.fill) { o.push(r.fill + ' RG'); top.stroke = r.fill; }
			var lw = num(r.size * 0.03);
			if (top.lw !== lw) { o.push(lw + ' w'); top.lw = lw; }
		}
		var tr = r.bold ? 2 : 0;
		var s = 'BT /' + r.font.res + ' ' + num5(r.size) + ' Tf ' + (top.tr !== tr ? tr + ' Tr ' : '') +
			[r.m[0], r.m[1], r.m[2], r.m[3]].map(num5).join(' ') + ' ' + num(r.x0) + ' ' + num(r.y0) + ' Tm [' + r.items.join('') + '] TJ ET';
		top.tr = tr;
		o.push(s);
	};

	Writer.prototype.glyph = function (f, gid, x, y, cps) {
		var st = this.st, pg = this.page, e = f.entry;
		var T = st.T, sizeMm = st.size * 25.4 / 72;
		var X = T[0] * x + T[2] * y + T[4], Y = T[1] * x + T[3] * y + T[5];
		var a = T[0], b = T[1], c = -T[2], d = -T[3];
		if (f.italic) { c = 0.2126 * a + c; d = 0.2126 * b + d; }   // synthetic italic: ~12 degrees, like FreeType's oblique
		var cid = e.cid(gid, cps);
		var fill = rgbStr(st.bType === 'grad' && st.grad && st.grad.stops.length ? st.grad.stops[0][1] : st.brush);
		var gsv = st.brushA < 255 ? this.gs(st.brushA / 255, st.brushA / 255) : '';
		// a viewer advances by the /W width (1/1000 em, rounded): position the next glyph from that, not from the font
		var r = pg.run, adv = e.widths.get(cid) / 1000 * sizeMm, dx = 0;
		if (r && r.font === e && r.size === sizeMm && r.fill === fill && r.gs === gsv && r.bold === f.bold &&
			r.m[0] === a && r.m[1] === b && r.m[2] === c && r.m[3] === d && b === 0 && c === 0 && a > 0 && Math.abs(Y - r.y) < 1e-4 &&
			(dx = (X - r.nextX) / a) < sizeMm && dx > -0.25 * sizeMm) {   // a large jump (tab, column) starts a new run
			var adj = -dx / sizeMm * 1000;
			if (Math.abs(adj) > 0.05) r.items.push(num(adj));
			r.items.push('<' + hex4(cid) + '>');
			r.nextX = X + adv * a;
		} else {
			this.flushRun();
			pg.res.F[e.res] = e.ref;
			pg.run = { font: e, size: sizeMm, fill: fill, gs: gsv, bold: f.bold, m: [a, b, c, d], x0: X, y0: Y, y: Y,
				nextX: X + adv * a, items: ['<' + hex4(cid) + '>'] };
		}
		this.stats.glyphs++;
	};

	Writer.prototype.emitPath = function (o, p) {
		for (var i = 0; i < p.length; i++) {
			var s = p[i];
			if (s[0] === 'm') o.push(num(s[1]) + ' ' + num(s[2]) + ' m');
			else if (s[0] === 'l') o.push(num(s[1]) + ' ' + num(s[2]) + ' l');
			else if (s[0] === 'c') o.push(num(s[1]) + ' ' + num(s[2]) + ' ' + num(s[3]) + ' ' + num(s[4]) + ' ' + num(s[5]) + ' ' + num(s[6]) + ' c');
			else o.push('h');
		}
	};
	function bboxOf(p) {
		var b = [1e9, 1e9, -1e9, -1e9];
		for (var i = 0; i < p.length; i++) for (var q = 1; q + 1 < p[i].length; q += 2) {
			b[0] = Math.min(b[0], p[i][q]); b[1] = Math.min(b[1], p[i][q + 1]);
			b[2] = Math.max(b[2], p[i][q]); b[3] = Math.max(b[3], p[i][q + 1]);
		}
		return b;
	}
	var JOIN = [0, 2, 1, 0, 0]; // Aggplus miter, bevel, round -> PDF 0 miter, 1 round, 2 bevel

	Writer.prototype.drawPath = async function (type) {
		var st = this.st, pg = this.page, path = pg.path, o = pg.out;
		pg.path = [];
		if (!path.length) return;
		this.flushRun();
		var fill = (type & 0x300) !== 0, stroke = (type & 1) !== 0, eo = (type & 0x200) !== 0;
		var special = fill && (st.bType === 3008 || st.bType === 'grad');
		this.stats.paths++;
		o.push('q');
		if (special) {
			if (st.bType === 'grad' && st.grad) {
				var sh = this.shading(st.grad, st.T);
				var a0 = st.grad.alpha;
				this.emitPath(o, path);
				o.push(eo ? 'W* n' : 'W n');
				if (a0 < 255) o.push(this.gs(a0 / 255, 1));
				if (sh) o.push('/' + sh + ' sh');
			} else if (st.tex) {
				var T = st.T, img, cm;
				if (st.rectOn && st.rect) {
					var R = st.rect, x = R[0], y = R[1], w = R[2], h = R[3];
					img = await this.image(st.tex, w, h);
					cm = [T[0] * w, T[1] * w, -T[2] * h, -T[3] * h, T[0] * x + T[2] * (y + h) + T[4], T[1] * x + T[3] * (y + h) + T[5]];
				} else {
					var bb = bboxOf(path);
					img = await this.image(st.tex, bb[2] - bb[0], bb[3] - bb[1]);
					cm = [bb[2] - bb[0], 0, 0, -(bb[3] - bb[1]), bb[0], bb[3]];
				}
				if (st.texMode === 1) this.skip('texture-tile');
				pg.res.X[img.res] = img.ref;
				this.emitPath(o, path);
				o.push(eo ? 'W* n' : 'W n');
				var ta = Math.min(st.texA, st.brushA);
				if (ta < 255) o.push(this.gs(ta / 255, 1));
				o.push(cm.map(num).join(' ') + ' cm /' + img.res + ' Do');
			}
			o.push('Q');
			if (!stroke) return;
			o.push('q');
			fill = false;
		}
		this.emitPath(o, path);
		var fa = fill ? st.brushA / 255 : 1, sa = stroke ? st.penA / 255 : 1;
		if (fa < 1 || sa < 1) o.push(this.gs(fa, sa));
		if (fill) o.push(rgbStr(st.brush) + ' rg');
		if (stroke) {
			o.push(rgbStr(st.pen) + ' RG ' + num(st.pw) + ' w');
			if (st.join) o.push(JOIN[st.join] + ' j');
			if (st.dash && st.dash.length) o.push('[' + st.dash.map(num).join(' ') + '] 0 d');
		}
		o.push(fill && stroke ? (eo ? 'B*' : 'B') : fill ? (eo ? 'f*' : 'f') : stroke ? 'S' : 'n');
		o.push('Q');
	};

	Writer.prototype.drawImage = async function (path, x, y, w, h) {
		var T = this.st.T, pg = this.page;
		this.flushRun();
		var img = await this.image(path, w, h);
		pg.res.X[img.res] = img.ref;
		var cm = [T[0] * w, T[1] * w, -T[2] * h, -T[3] * h, T[0] * x + T[2] * (y + h) + T[4], T[1] * x + T[3] * (y + h) + T[5]];
		pg.out.push('q ' + cm.map(num).join(' ') + ' cm /' + img.res + ' Do Q');
	};

	Writer.prototype.text = async function (str, x, y) {
		var f = await this.font(this.st.font, this.st.style), sizeMm = this.st.size * 25.4 / 72;
		var cx = x;
		for (var i = 0; i < str.length; i++) {
			var cp = str.codePointAt(i);
			if (cp > 0xFFFF) i++;
			var gid = f.entry.gidForCodePoint(cp);
			this.glyph(f, gid, cx, y, [cp]);
			cx += f.entry.advance(gid) * sizeMm;
		}
		this.stats.texts++;
	};

	Writer.prototype.endPage = async function () {
		var pg = this.page;
		if (!pg) return;
		this.flushRun();
		this.page = null;
		var H = pg.h * MM;
		var head = 'q ' + num5(MM) + ' 0 0 ' + num5(-MM) + ' 0 ' + num(H) + ' cm\n';
		var body = latin1(head + pg.out.join('\n') + '\nQ\n');
		pg.out = null;
		pg.ref = this.pageRef(pg.index);
		pg.contentRef = this.alloc();
		await this.putStreamDeflate(pg.contentRef, '', body);
		this.pages.push(pg);
		this.stats.pages++;
	};

	// ---- decoder ---------------------------------------------------------------------------------
	var NEEDS_PAGE = {};
	[80, 83, 91, 92, 94, 97, 98, 99, 100, 111, 121, 122, 160, 161, 200, 201].forEach(function (c) { NEEDS_PAGE[c] = true; });
	Writer.prototype.feed = async function (u8) {
		var r = new Reader(u8), st = this.st;
		var t0 = (typeof performance !== 'undefined' ? performance.now() : Date.now());
		var tp = function (x, y) { var T = st.T; return [T[0] * x + T[2] * y + T[4], T[1] * x + T[3] * y + T[5]]; };
		while (r.more()) {
			var c = r.byte(), pg = this.page;
			if (!pg && NEEDS_PAGE[c]) pg = this.page = new Page(this, this.stats.pages); // defensive: drawing outside a page
			switch (c) {
				case 202:
					// sdkjs starts every page with a fresh CMetafile (pen, brush, font, transform): reset like it
					if (this.page) await this.endPage();
					this.page = new Page(this, this.nextPageIndex !== undefined ? this.nextPageIndex : this.stats.pages);
					this.nextPageIndex = undefined;
					this.st = st = this.freshState();
					break;
				case 200: this.page.w = r.dbl(); break;
				case 201: this.page.h = r.dbl(); break;
				case 203: await this.endPage(); break;
				// pen
				case 1: st.pen = r.int(); break;
				case 2: st.penA = r.byte(); break;
				case 3: st.pw = r.dbl(); break;
				case 4: {
					var ds = r.byte(), d = null, pw = Math.max(st.pw, 0.1);
					if (ds === 5) { var nd = r.int(); d = []; for (var i = 0; i < nd; i++) d.push(r.dbl()); }
					else if (ds === 1) d = [3 * pw, pw];
					else if (ds === 2) d = [pw, pw];
					else if (ds === 3) d = [3 * pw, pw, pw, pw];
					else if (ds === 4) d = [3 * pw, pw, pw, pw, pw, pw];
					st.dash = d; break;
				}
				case 5: case 6: r.byte(); break;  // line caps (not written by sdkjs today)
				case 7: st.join = r.byte(); break;
				// brush
				case 21: st.bType = r.int(); st.grad = null; break;
				case 22: st.brush = r.int(); break;
				case 23: r.int(); break;
				case 24: st.brushA = r.byte(); break;
				case 25: r.byte(); break;
				case 26: st.tex = r.str(); break;
				case 27: st.texA = r.byte(); break;
				case 28: st.texMode = r.byte(); break;
				case 29: st.rect = [r.dbl(), r.dbl(), r.dbl(), r.dbl()]; break;
				case 30: st.rectOn = !!r.byte(); if (!st.rectOn) st.rect = null; break;
				case 31: {
					r.byte();
					var g = { type: 'lin', p: [0, 0, 0, 0], stops: [], alpha: 255 };
					for (;;) {
						var sc = r.byte();
						if (sc === 251) break;
						if (sc === 0) { r.p += 5; g.type = 'lin'; g.p = [r.dbl(), r.dbl(), r.dbl(), r.dbl()]; }
						else if (sc === 1) { r.p += 1; g.type = 'rad'; g.p = [r.dbl(), r.dbl(), r.dbl(), r.dbl(), r.dbl(), r.dbl()]; }
						else if (sc === 2) {
							var nc = r.int(), amin = 255;
							for (var q = 0; q < nc; q++) { var pos = r.dbl(), col = r.uint(); g.stops.push([pos, col]); amin = Math.min(amin, col >>> 24); }
							g.alpha = amin;
						}
						else throw Unsupported('gradient', 'sub-command ' + sc);
					}
					st.grad = g; st.bType = 'grad';
					break;
				}
				case 32: st.tex = r.str2(); break;
				case 33: this.skip('reset-rotation'); break;
				// font
				case 41: st.font = r.str(); break;
				case 42: st.size = r.dbl(); break;
				case 43: st.style = r.int(); break;
				// text
				case 80: { var s = r.str(), tx = r.dbl(), ty = r.dbl(); await this.text(s, tx, ty); break; }
				case 83: {
					var gid = r.uint(), gx = r.dbl(), gy = r.dbl(), nu = r.int(), cps = [];
					for (var u = 0; u < nu; u++) cps.push(r.int());
					var f = this.fontByName.get(st.font + '|' + st.style) || await this.font(st.font, st.style);
					this.glyph(f, gid, gx, gy, cps);
					break;
				}
				// paths
				case 100: pg.path = []; break;
				case 91: { var a = tp(r.dbl(), r.dbl()); pg.path.push(['m', a[0], a[1]]); break; }
				case 92: { var b = tp(r.dbl(), r.dbl()); pg.path.push(['l', b[0], b[1]]); break; }
				case 94: {
					var A = tp(r.dbl(), r.dbl()), B = tp(r.dbl(), r.dbl()), C = tp(r.dbl(), r.dbl());
					pg.path.push(['c', A[0], A[1], B[0], B[1], C[0], C[1]]); break;
				}
				case 97: pg.path.push(['z']); break;
				case 98: if (!pg.clipping) pg.path = []; break;
				case 99: await this.drawPath(r.int()); break;
				case 104: case 105: r.dbl(); r.dbl(); break;
				// images
				case 111: { var ip = r.str2(), ix = r.dbl(), iy = r.dbl(), iw = r.dbl(), ih = r.dbl(); await this.drawImage(ip, ix, iy, iw, ih); break; }
				// commands
				case 121: {
					var bc = r.int();
					if (bc === 0x20) { pg.path = []; pg.clipping = true; }
					else if (bc === 0x40) {
						this.flushRun();
						pg.out.push('Q q ' + num5(MM) + ' 0 0 ' + num5(-MM) + ' 0 ' + num(pg.h * MM) + ' cm');
						pg.top = { fill: '', gs: '', tr: 0, lw: '' };
					}
					break;
				}
				case 122: {
					var ec = r.int();
					if (ec === 0x20 && pg.clipping) {
						this.flushRun();
						if (pg.path.length) { this.emitPath(pg.out, pg.path); pg.out.push('W n'); }
						pg.clipping = false; pg.path = [];
					}
					break;
				}
				case 130: st.T = [r.dbl(), r.dbl(), r.dbl(), r.dbl(), r.dbl(), r.dbl()]; break;
				// links and document structure
				case 160: {
					var hx = r.dbl(), hy = r.dbl(), hw = r.dbl(), hh = r.dbl(), url = r.str(); r.str();
					pg.annots.push({ uri: url, rect: [tp(hx, hy), tp(hx + hw, hy + hh), tp(hx, hy + hh), tp(hx + hw, hy)] });
					this.stats.links++;
					break;
				}
				case 161: {
					var lx = r.dbl(), ly = r.dbl(), lw = r.dbl(), lh = r.dbl(), dx = r.dbl(), dy = r.dbl(), dp = r.int();
					pg.annots.push({ page: dp, dx: dx, dy: dy, rect: [tp(lx, ly), tp(lx + lw, ly + lh), tp(lx, ly + lh), tp(lx + lw, ly)] });
					this.stats.links++;
					break;
				}
				case 163: {
					var fl = r.int(), keys = ['title', 'author', 'subject', 'keywords'];
					for (var k = 0; k < 4; k++) if (fl & (1 << k)) this.info[keys[k]] = r.str();
					break;
				}
				case 169: {
					var start169 = r.p, len169 = r.int(), nh = r.int(), hs = [];
					for (var hI = 0; hI < nh; hI++) hs.push({ lvl: r.int(), page: r.int(), x: r.dbl(), y: r.dbl(), text: r.str() });
					r.p = start169 + len169;
					this.headings = hs;
					break;
				}
				case 162: throw Unsupported('form-fields');
				case 164: case 165: case 166: case 170: case 207: case 208: {
					var st0 = r.p, ln = r.int(); r.p = st0 + ln; this.skip('cmd' + c); break;
				}
				default: {
					if (c >= 160) { var s0 = r.p, l0 = r.int(); r.p = s0 + l0; this.skip('cmd' + c); break; }
					throw Unsupported('command', String(c));
				}
			}
		}
		this.stats.ms += (typeof performance !== 'undefined' ? performance.now() : Date.now()) - t0;
	};

	// ---- finish ----------------------------------------------------------------------------------
	function annotRect(pts, H) {
		var xs = pts.map(function (p) { return p[0] * MM; }), ys = pts.map(function (p) { return H - p[1] * MM; });
		return [Math.min.apply(null, xs), Math.min.apply(null, ys), Math.max.apply(null, xs), Math.max.apply(null, ys)].map(num).join(' ');
	}

	Writer.prototype.finish = async function () {
		var t0 = (typeof performance !== 'undefined' ? performance.now() : Date.now());
		if (this.page) await this.endPage();
		if (!this.pages.length) throw Unsupported('empty');
		var me = this, env = this.env;
		var pagesRef = this.alloc();
		var fontsInfo = [];
		for (var fe of this.fonts.values()) fontsInfo.push(await fe.write());
		this.stats.fonts = fontsInfo;

		var byIndex = {};
		this.pages.forEach(function (p) { byIndex[p.index] = p; });
		function dest(pi, x, y) {
			var p = byIndex[pi];
			if (!p) return null;
			return '[' + p.ref + ' 0 R /XYZ ' + num(x * MM) + ' ' + num((p.h - y) * MM) + ' 0]';
		}
		this.pages.forEach(function (p) {
			var H = p.h * MM, annots = [];
			p.annots.forEach(function (a) {
				var ar = me.alloc(), action;
				if (a.uri !== undefined) {
					action = '/A << /Type /Action /S /URI /URI ' + textString(a.uri) + ' >>';
				} else {
					var d = dest(a.page, a.dx, a.dy);
					if (!d) return;
					action = '/Dest ' + d;
				}
				me.put(ar, '<< /Type /Annot /Subtype /Link /Rect [' + annotRect(a.rect, H) + '] /Border [0 0 0] ' + action + ' >>');
				annots.push(ar + ' 0 R');
			});
			var res = [];
			var dictOf = function (o) { return Object.keys(o).map(function (k) { return '/' + k + ' ' + o[k] + ' 0 R'; }).join(' '); };
			if (Object.keys(p.res.F).length) res.push('/Font << ' + dictOf(p.res.F) + ' >>');
			if (Object.keys(p.res.X).length) res.push('/XObject << ' + dictOf(p.res.X) + ' >>');
			if (Object.keys(p.res.G).length) res.push('/ExtGState << ' + dictOf(p.res.G) + ' >>');
			if (Object.keys(p.res.S).length) res.push('/Shading << ' + dictOf(p.res.S) + ' >>');
			me.put(p.ref, '<< /Type /Page /Parent ' + pagesRef + ' 0 R /MediaBox [0 0 ' + num(p.w * MM) + ' ' + num(H) + ']' +
				' /Resources << ' + res.join(' ') + ' >> /Contents ' + p.contentRef + ' 0 R' +
				(annots.length ? ' /Annots [' + annots.join(' ') + ']' : '') + ' >>');
		});
		// pages referenced by links but never written (out of the printed range): drop the refs
		this.pageRefs.forEach(function (ref) { if (!me.objs[ref]) me.put(ref, 'null'); });
		this.put(pagesRef, '<< /Type /Pages /Kids [' + this.pages.map(function (p) { return p.ref + ' 0 R'; }).join(' ') + '] /Count ' + this.pages.length + ' >>');

		// outline (document headings)
		var outlinesRef = 0;
		if (this.headings && this.headings.length) {
			var items = this.headings.filter(function (h) { return byIndex[h.page]; }).map(function (h) { return { h: h, ref: me.alloc(), kids: [] }; });
			if (items.length) {
				outlinesRef = this.alloc();
				var rootItem = { ref: outlinesRef, kids: [], lvl: -1 }, stack = [rootItem];
				items.forEach(function (it) {
					while (stack.length > 1 && stack[stack.length - 1].h.lvl >= it.h.lvl) stack.pop();
					it.parent = stack[stack.length - 1];
					it.parent.kids.push(it);
					stack.push(it);
				});
				var count = function (it) { var n = it.kids.length; it.kids.forEach(function (k) { n += count(k); }); return n; };
				items.forEach(function (it) {
					var sib = it.parent.kids, i = sib.indexOf(it);
					me.put(it.ref, '<< /Title ' + textString(it.h.text || '') + ' /Parent ' + it.parent.ref + ' 0 R' +
						(i > 0 ? ' /Prev ' + sib[i - 1].ref + ' 0 R' : '') + (i + 1 < sib.length ? ' /Next ' + sib[i + 1].ref + ' 0 R' : '') +
						(it.kids.length ? ' /First ' + it.kids[0].ref + ' 0 R /Last ' + it.kids[it.kids.length - 1].ref + ' 0 R /Count -' + count(it) : '') +
						' /Dest ' + dest(it.h.page, it.h.x, it.h.y) + ' >>');
				});
				this.put(outlinesRef, '<< /Type /Outlines /First ' + rootItem.kids[0].ref + ' 0 R /Last ' + rootItem.kids[rootItem.kids.length - 1].ref +
					' 0 R /Count ' + items.length + ' >>');
			}
		}

		var catalogRef = this.alloc();
		var lang = env.lang ? ' /Lang ' + textString(env.lang) : '';
		this.put(catalogRef, '<< /Type /Catalog /Pages ' + pagesRef + ' 0 R' + (outlinesRef ? ' /Outlines ' + outlinesRef + ' 0 R' : '') + lang + ' >>');
		var infoRef = this.alloc(), info = this.info, now = pdfDate(new Date());
		var title = info.title || env.title || '';
		this.put(infoRef, '<< ' + (title ? '/Title ' + textString(title) + ' ' : '') + (info.author ? '/Author ' + textString(info.author) + ' ' : '') +
			(info.subject ? '/Subject ' + textString(info.subject) + ' ' : '') + (info.keywords ? '/Keywords ' + textString(info.keywords) + ' ' : '') +
			'/Creator ' + textString(env.creator || 'ElixirOffice') + ' /Producer ' + textString(env.producer || 'ElixirOffice') +
			' /CreationDate (' + now + ') /ModDate (' + now + ') >>');

		// serialize
		var parts = [latin1('%PDF-1.7\n%âãÏÓ\n')], pos = parts[0].length, offsets = [0];
		for (var id = 1; id < this.objs.length; id++) {
			var ob = this.objs[id] || latin1(id + ' 0 obj\nnull\nendobj\n');
			offsets[id] = pos;
			parts.push(ob);
			pos += ob.length;
		}
		var xref = ['xref', '0 ' + this.objs.length, '0000000000 65535 f '];
		for (var j = 1; j < this.objs.length; j++) xref.push(('0000000000' + offsets[j]).slice(-10) + ' 00000 n ');
		var idh = '';
		for (var b = 0; b < 16; b++) idh += hex4(Math.floor(Math.random() * 256)).slice(2);
		parts.push(latin1(xref.join('\n') + '\ntrailer\n<< /Size ' + this.objs.length + ' /Root ' + catalogRef + ' 0 R /Info ' + infoRef +
			' 0 R /ID [<' + idh + '> <' + idh + '>] >>\nstartxref\n' + pos + '\n%%EOF\n'));
		this.objs = null;
		this.stats.ms += (typeof performance !== 'undefined' ? performance.now() : Date.now()) - t0;
		return concat(parts);
	};

	root.EOPdf = { Writer: Writer, Unsupported: Unsupported, version: 1 };
})(typeof window !== 'undefined' ? window : globalThis);
