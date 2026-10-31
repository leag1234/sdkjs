// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

function scanner(options) {
    const source = fs.readFileSync(path.join(root, 'common/clientpdf/wasm-pdf.js'), 'utf8');
    const context = {
        URL,
        AscCommon: {ClientPdf: {
            isCjkCodePoint(codepoint) {
                return (codepoint >= 0x2E80 && codepoint <= 0xA4CF)
                    || (codepoint >= 0xAC00 && codepoint <= 0xD7FF)
                    || (codepoint >= 0xF900 && codepoint <= 0xFAFF)
                    || (codepoint >= 0x20000 && codepoint <= 0x323AF);
            },
            unsupported(code, detail) {
                return Object.assign(new Error(code), {eoUnsupported: true, code, detail});
            }
        }},
        location: {href: 'https://example.invalid/sdkjs/common/clientpdf/wasm-pdf.js'}
    };
    context.window = context;
    vm.runInNewContext(source, context);
    return new context.EOWasmPdf.Scanner(options);
}

function textCommand(value) {
    const bytes = new Uint8Array(1 + 2 + value.length * 2 + 8);
    const view = new DataView(bytes.buffer);
    bytes[0] = 80;
    view.setUint16(1, value.length, true);
    for (let i = 0; i < value.length; i++)
        view.setUint16(3 + i * 2, value.charCodeAt(i), true);
    return bytes;
}

function glyphCommand(gid, codepoints) {
    const bytes = new Uint8Array(1 + 12 + 4 + codepoints.length * 4);
    const view = new DataView(bytes.buffer);
    bytes[0] = 83;
    view.setUint32(1, gid, true);
    view.setInt32(13, codepoints.length, true);
    codepoints.forEach((codepoint, index) => view.setInt32(17 + index * 4, codepoint, true));
    return bytes;
}

test('WASM PDF scanner fails closed instead of emitting CJK with a stale glyph id', () => {
    assert.throws(() => scanner().feed(textCommand('\u4E2D')), {code: 'cjk'});
    assert.throws(() => scanner().feed(glyphCommand(0, [0x4E2D])), {code: 'cjk'});
});

test('WASM PDF scanner still accepts non-CJK text and glyph commands', () => {
    assert.doesNotThrow(() => scanner().feed(textCommand('A')));
    assert.doesNotThrow(() => scanner().feed(glyphCommand(36, [0x41])));
});

test('WASM PDF scanner requires an explicit boolean CJK remap capability', () => {
    for (const options of [undefined, {}, {nativeCjkGlyphRemap: false}, {nativeCjkGlyphRemap: 'true'}, {nativeCjkGlyphRemap: 1}]) {
        assert.throws(() => scanner(options).feed(textCommand('\u4E2D')), {code: 'cjk'});
        assert.throws(() => scanner(options).feed(glyphCommand(0, [0x4E2D])), {code: 'cjk'});
    }
});

test('WASM PDF scanner accepts only single BMP CJK scalars with remap capability', () => {
    assert.doesNotThrow(() => scanner({nativeCjkGlyphRemap: true}).feed(textCommand('\u4E2D')));
    assert.doesNotThrow(() => scanner({nativeCjkGlyphRemap: true}).feed(glyphCommand(0, [0x4E2D])));
});

test('WASM PDF scanner rejects supplementary CJK and variation selectors with remap capability', () => {
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(textCommand('\uD840\uDC00')), {code: 'cjk'});
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(glyphCommand(0, [0x20000])), {code: 'cjk'});
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(textCommand('\u4E2D\uFE0F')), {code: 'cjk'});
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(textCommand('\u4E2D\u{E0100}')), {code: 'cjk'});
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(textCommand('\u4E2D\u200D')), {code: 'cjk'});
});

test('WASM PDF scanner rejects multi-scalar glyph runs with remap capability', () => {
    assert.throws(() => scanner({nativeCjkGlyphRemap: true}).feed(glyphCommand(0, [0x4E2D, 0x4E2D])), {code: 'cjk'});
});

test('WASM PDF scanner accepts non-CJK in every remap capability mode', () => {
    for (const options of [undefined, {}, {nativeCjkGlyphRemap: false}, {nativeCjkGlyphRemap: 'true'}, {nativeCjkGlyphRemap: 1}, {nativeCjkGlyphRemap: true}]) {
        assert.doesNotThrow(() => scanner(options).feed(textCommand('A')));
        assert.doesNotThrow(() => scanner(options).feed(glyphCommand(36, [0x41])));
    }
});

test('WASM PDF scanner keeps joiners and variation selectors in non-CJK text accepted', () => {
    for (const options of [undefined, {nativeCjkGlyphRemap: true}]) {
        assert.doesNotThrow(() => scanner(options).feed(textCommand('A\u200DB')));
        assert.doesNotThrow(() => scanner(options).feed(textCommand('\u00A9\uFE0F')));
    }
});

test('WASM PDF writer enables CJK remap only for a strict true capability from the runtime', () => {
    const src = fs.readFileSync(path.join(root, 'common/clientpdf/wasm-pdf.js'), 'utf8');
    assert.match(src, /nativeCjkGlyphRemap: env\['wasmCjkGidRemap'\] === true/);
    const cp = fs.readFileSync(path.join(root, 'common/clientpdf/clientpdf.js'), 'utf8');
    assert.match(cp, /manifest\.capabilities\.cjk_gid_remap_bmp === 1/);
    assert.match(cp, /env\['wasmCjkGidRemap'\] = cjkGidRemap/);
});

test.todo('native PDF replay selects the fallback font and remaps CJK glyph ids before conversion');
