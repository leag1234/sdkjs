// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../../common/editorscommon.js'), 'utf8');

function extractFunction(name)
{
	const marker = 'function ' + name + '(';
	const start = source.indexOf(marker);
	assert(start >= 0, name);
	const brace = source.indexOf('{', start);
	let depth = 0;
	for (let i = brace; i < source.length; ++i)
	{
		if (source[i] === '{') ++depth;
		else if (source[i] === '}' && --depth === 0)
			return source.slice(start, i + 1);
	}
	assert.fail('Unterminated ' + name);
}

function makeEnv(opts)
{
	opts = opts || {};
	const clicks = [];
	const warnings = [];
	const listeners = {document: {}, window: {}};
	const timers = [];
	function addListener(map, type, fn) { (map[type] = map[type] || []).push(fn); }
	function removeListener(map, type, fn) {
		const arr = map[type] || [];
		const i = arr.indexOf(fn);
		if (i >= 0) arr.splice(i, 1);
	}
	function fire(map, type, event) { (map[type] || []).slice().forEach(fn => fn(event)); }
	const documentStub = {
		createElement: () => {
			const el = {
				style: {},
				classList: {add: () => {}},
				setAttribute: function (k, v) { this[k] = v; },
				click: function () {
					clicks.push(this.href);
					if (opts.fireViolation) {
						const ev = {blockedURI: this.href};
						queueMicrotask(() => {
							fire(listeners.document, 'securitypolicyviolation', ev);
							fire(listeners.window, 'securitypolicyviolation', ev);
						});
					}
					if (opts.throwOnClick) throw new Error('click blocked');
				}
			};
			return el;
		},
		body: {appendChild: () => {}, removeChild: () => {}},
		addEventListener: (t, f) => addListener(listeners.document, t, f),
		removeEventListener: (t, f) => removeListener(listeners.document, t, f)
	};
	const windowStub = {
		navigator: {},
		console: null,
		URL: {
			createObjectURL: () => 'blob:mock-url',
			revokeObjectURL: () => {}
		},
		addEventListener: (t, f) => addListener(listeners.window, t, f),
		removeEventListener: (t, f) => removeListener(listeners.window, t, f)
	};
	const context = vm.createContext({
		window: windowStub,
		document: documentStub,
		Blob: function (parts, o) { this.parts = parts; this.type = o && o.type; },
		setTimeout: (fn, ms) => { timers.push({fn, ms}); return timers.length; },
		clearTimeout: () => {},
		console: {warn: (msg) => warnings.push(msg), error: () => {}},
		AscCommon: {}
	});
	context.window.AscCommon = context.AscCommon;
	windowStub.console = context.console;
	const bundle = [extractFunction('DownloadFileFromBytes'), extractFunction('downloadViaBlobLink')].join('\n');
	const fn = vm.runInContext('(function(){' + bundle + '; return DownloadFileFromBytes;})()', context);
	return {fn, clicks, timers, warnings, context};
}

test('working blob download: exactly one click with a blob: href, callback not called', () =>
{
	const env = makeEnv();
	let calls = 0;
	env.fn.call(env.context, new Uint8Array([1, 2, 3]), 'doc.pdf', 'application/pdf', undefined, () => { ++calls; });
	assert.equal(env.clicks.length, 1);
	assert.equal(env.clicks[0], 'blob:mock-url');
	assert.equal(calls, 0);
});

test('CSP-blocked blob navigation: callback called exactly once, no second link click', async () =>
{
	const env = makeEnv({fireViolation: true});
	let calls = 0;
	env.fn.call(env.context, new Uint8Array([37, 80, 68, 70]), 'doc.pdf', 'application/pdf', undefined, () => { ++calls; });
	await new Promise(resolve => setImmediate(resolve));
	assert.equal(env.clicks.length, 1);
	assert.equal(env.clicks[0], 'blob:mock-url');
	assert.equal(calls, 1);
});

test('blob click throws: callback called exactly once', () =>
{
	const env = makeEnv({throwOnClick: true});
	let calls = 0;
	env.fn.call(env.context, new Uint8Array([1]), 'doc.pdf', 'application/pdf', undefined, () => { ++calls; });
	assert.equal(calls, 1);
});

test('no callback: blocked blob download does not throw, only one console.warn', async () =>
{
	const env = makeEnv({fireViolation: true});
	env.fn.call(env.context, new Uint8Array([1, 2, 3]), 'doc.pdf', 'application/pdf');
	await new Promise(resolve => setImmediate(resolve));
	assert.equal(env.clicks.length, 1);
	assert.equal(env.warnings.length, 1);
});
