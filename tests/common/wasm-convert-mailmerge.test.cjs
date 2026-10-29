// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

test('Wasm converter exposes the core mail-merge JSON format', () => {
    const context = {
        document: {currentScript: {src: 'https://local.invalid/common/wasm-convert/api.js'}},
        URL
    };
    context.window = context;
    vm.runInNewContext(fs.readFileSync(path.resolve(__dirname,
        '../../common/wasm-convert/api.js'), 'utf8'), context);
    assert.equal(context.EOWasmConvert.formats.json, 0x0808);
    const worker = fs.readFileSync(path.resolve(__dirname,
        '../../common/wasm-convert/worker.js'), 'utf8');
    assert.match(worker, /json:2056/);
});
