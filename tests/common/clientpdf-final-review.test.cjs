// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');
const source = fs.readFileSync(path.join(root, 'common/clientpdf/clientpdf.js'), 'utf8');
const modes = {Edit: 0, Final: 1, Original: 2, Simple: 3};

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

function loadClientPdf(createWriter) {
    const AscCommon = {
        CDocumentRenderer: Renderer,
        CMemory: Memory,
        g_oTextMeasurer: {m_oManager: {}},
        g_oDocumentUrls: {}
    };
    const window = {
        AscCommon,
        AscWord: {},
        localStorage: {getItem: () => null, setItem() {}},
        location: {href: 'https://local.invalid/editor/'},
        performance: {now: () => Date.now()}
    };
    const context = {
        window,
        AscCommon,
        AscFonts: {g_fonts_streams: [], g_map_font_index: {}, g_font_infos: []},
        Asc: {c_oAscDisplayModeInReview: modes},
        Uint8Array,
        ArrayBuffer,
        Buffer,
        Promise,
        Date,
        URL,
        Blob,
        performance: window.performance,
        setTimeout,
        clearTimeout,
        setInterval,
        clearInterval,
        console
    };
    vm.runInNewContext(source, context);
    const client = window.AscCommon.ClientPdf;
    client.registerBackend({id: 'test', load: async () => {}, create: createWriter});
    client.setBackend('test');
    client.__testWindow = window;
    return client;
}

function fixture(priorMode, options = {}) {
    const revisionItems = [
        {text: 'stable', final: true, original: true},
        {text: 'inserted', final: true, original: false},
        {text: 'deleted', final: false, original: true},
        {text: 'move-from', final: false, original: true},
        {text: 'move-to', final: true, original: false}
    ];
    const priorReview = priorMode === modes.Final || priorMode === modes.Original;
    const initial = {
        mode: priorMode,
        fast: priorReview ? false : true,
        lock: priorReview,
        dirty: true,
        undo: ['u1', 'u2'],
        redo: ['r1'],
        historyIdentity: {id: 'history-object'},
        sync: 'caller-sentinel'
    };
    const state = {...initial, undo: initial.undo.slice(), redo: initial.redo.slice()};
    const transitions = [];
    const draws = [];
    const feeds = [];
    const signal = options.signal || {aborted: false};

    function visible() {
        if (state.mode === modes.Final) return revisionItems.filter(item => item.final).map(item => item.text);
        if (state.mode === modes.Original) return revisionItems.filter(item => item.original).map(item => item.text);
        return revisionItems.map(item => item.text);
    }
    const logic = {
        GetDisplayModeInReview: () => state.mode,
        SetDisplayModeInReview(mode) {
            transitions.push([state.mode, mode]);
            state.mode = mode;
            state.lock = mode === modes.Final || mode === modes.Original;
            state.fast = !state.lock;
        },
        DrawPage(page, renderer) {
            draws.push({
                page,
                mode: state.mode,
                lock: state.lock,
                sync: client.__testWindow.NATIVE_EDITOR_ENJINE_SYNC_RECALC,
                visible: visible()
            });
            if (options.abortDuringDraw) signal.aborted = true;
            if (options.throwDuringDraw) throw new Error('draw failed');
            renderer.Memory.write(visible().join('|'));
        }
    };
    const drawing = {
        m_lPagesCount: 2,
        m_lCountCalculatePages: 2,
        m_arrPages: [{width_mm: 210, height_mm: 297}, {width_mm: 210, height_mm: 297}],
        m_oLogicDocument: logic
    };
    const api = {
        WordControl: {m_oDrawingDocument: drawing, m_oLogicDocument: logic},
        documentTitle: 'tracked.docx',
        ShowParaMarks: true,
        asc_getCoreProps: () => ({}),
        asc_isFinalizedVersion: () => false
    };
    const client = loadClientPdf(() => ({
        stats: {},
        nextPageIndex: 0,
        async feed(bytes) {
            feeds.push({value: Buffer.from(bytes).toString(), mode: state.mode, lock: state.lock});
            if (options.throwDuringFeed && feeds.length > 1) throw new Error('feed failed');
        },
        async finish() { return Uint8Array.from([1, 2, 3]); }
    }));
    client.__testWindow.NATIVE_EDITOR_ENJINE_SYNC_RECALC = initial.sync;
    return {api, client, state, initial, transitions, draws, feeds, signal};
}

function assertRestored(value) {
    assert.equal(value.state.mode, value.initial.mode);
    assert.equal(value.state.fast, value.initial.fast);
    assert.equal(value.state.lock, value.initial.lock);
    assert.equal(value.state.dirty, value.initial.dirty);
    assert.deepEqual(value.state.undo, value.initial.undo);
    assert.deepEqual(value.state.redo, value.initial.redo);
    assert.equal(value.state.historyIdentity, value.initial.historyIdentity);
    assert.equal(value.client.__testWindow.NATIVE_EDITOR_ENJINE_SYNC_RECALC, value.initial.sync);
}

for (const prior of [modes.Edit, modes.Simple, modes.Original, modes.Final]) {
    test('PDF captures final revisions and restores prior review state ' + prior, async () => {
        const value = fixture(prior);
        await value.client.build(value.api, {signal: value.signal});
        assert.equal(value.draws.length, 2);
        for (const draw of value.draws) {
            assert.equal(draw.mode, modes.Final);
            assert.equal(draw.lock, true);
            assert.equal(draw.sync, prior === modes.Final ? value.initial.sync : true);
            assert.deepEqual(draw.visible, ['stable', 'inserted', 'move-to']);
        }
        for (const feed of value.feeds) assert.equal(feed.mode, prior);
        assert(value.feeds.some(feed => feed.value === 'stable|inserted|move-to'));
        assertRestored(value);
    });
}

for (const failure of ['draw', 'feed', 'cancel-during-draw']) {
    test('PDF restores state after ' + failure, async () => {
        const value = fixture(modes.Simple, {
            throwDuringDraw: failure === 'draw',
            throwDuringFeed: failure === 'feed',
            abortDuringDraw: failure === 'cancel-during-draw'
        });
        await assert.rejects(
            value.client.build(value.api, {signal: value.signal}),
            failure === 'cancel-during-draw' ? /cancelled/ : new RegExp(failure)
        );
        assert(value.draws.every(draw => draw.mode === modes.Final && draw.lock));
        assertRestored(value);
    });
}

test('pre-cancelled PDF never enters final review view', async () => {
    const value = fixture(modes.Edit, {signal: {aborted: true}});
    await assert.rejects(value.client.build(value.api, {signal: value.signal}), /cancelled/);
    assert.deepEqual(value.transitions, []);
    assert.deepEqual(value.draws, []);
    assertRestored(value);
});
