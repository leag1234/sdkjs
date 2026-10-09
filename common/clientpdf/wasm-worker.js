/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 * One fresh module per export: native globals never survive between jobs.
 */
'use strict';
self.onmessage = async function (event) {
	try {
		var job = event.data, start = performance.now();
		importScripts(job.runtime + 'x2t.js');
		var response = await fetch(job.runtime + 'x2t.wasm', { credentials: 'same-origin', cache: 'force-cache' });
		if (!response.ok) throw new Error('cannot load wasm module');
		var compiled = await WebAssembly.compile(await response.arrayBuffer());
		var mod = await createX2t({
			noInitialRun: true, print: function () {}, printErr: function () {},
			instantiateWasm: function (imports, receive) {
				var instance = new WebAssembly.Instance(compiled, imports);
				receive(instance, compiled);
				return instance.exports;
			}
		});
		var FS = mod.FS;
		['/p/media', '/fonts', '/tmp/pdf'].forEach(function (p) { FS.mkdirTree(p); });
		FS.writeFile('/p/in.bin', job.stream);
		job.files.forEach(function (f) { FS.writeFile(f.path, f.bytes); });
		FS.writeFile('/task.xml', '<?xml version="1.0" encoding="utf-8"?><TaskQueueDataConvert>' +
			'<m_sFileFrom>/p/in.bin</m_sFileFrom><m_sFileTo>/p/out.pdf</m_sFileTo>' +
			'<m_nFormatFrom>8196</m_nFormatFrom><m_nFormatTo>513</m_nFormatTo>' +
			'<m_bIsNoBase64>true</m_bIsNoBase64><m_sFontDir>/fonts</m_sFontDir>' +
			'<m_sTempDir>/tmp/pdf</m_sTempDir></TaskQueueDataConvert>');
		var initMs = performance.now() - start, at = performance.now(), rc;
		try { rc = mod.callMain(['/task.xml']); } catch (e) { if (typeof e.status === 'number') rc = e.status; else throw e; }
		if (rc !== 0) throw new Error('native PDF conversion failed: ' + rc);
		var bytes = FS.readFile('/p/out.pdf');
		if (bytes.length < 8 || String.fromCharCode.apply(null, bytes.subarray(0, 5)) !== '%PDF-') throw new Error('invalid PDF output');
		self.postMessage({ bytes: bytes, initMs: Math.round(initMs), convertMs: Math.round(performance.now() - at), heapBytes: mod.HEAPU8.length }, [bytes.buffer]);
	} catch (e) { self.postMessage({ error: String(e.message || e), code: 'wasm' }); }
};
