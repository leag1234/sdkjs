// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '../..');

test('mail-merge PDF uses an isolated synchronous layout and always restores editor state', () => {
    const drawing = fs.readFileSync(path.join(root, 'word/Drawing/DrawingDocument.js'), 'utf8');
    const pdf = fs.readFileSync(path.join(root, 'common/clientpdf/clientpdf.js'), 'utf8');
    const start = drawing.indexOf('this.GenerateMailMergePrint = function');
    const end = drawing.indexOf('this.GenerateSelectionPrint = function', start);
    assert(start > 0 && end > start);
    const method = drawing.slice(start, end);
    assert.match(method, /NATIVE_EDITOR_ENJINE_SYNC_RECALC/);
    assert.match(method, /logicDocument\.RecalculateFromStart\(false\)/);
    assert.match(method, /finally[\s\S]*m_oLogicDocument = sourceLogic/);
    assert.match(method, /m_oDrawingDocument = sourceDrawing/);
    assert.match(method, /logicDocument\.DrawingDocument = documentDrawing/);
    assert.match(pdf, /GenerateMailMergePrint\(options\.oDocumentMailMerge\)/);
    assert.match(pdf, /finally[\s\S]*if \(mergedPrint\)[\s\S]*dd\.printedDocument = null/);
});

test('mail-merge layout restores history, table registry and editor documents after success or failure', () => {
    const drawing = fs.readFileSync(path.join(root, 'word/Drawing/DrawingDocument.js'), 'utf8');
    const start = drawing.indexOf('this.GenerateMailMergePrint = function');
    const end = drawing.indexOf('this.GenerateSelectionPrint = function', start);
    const method = drawing.slice(start, end);
    const expression = method.slice(method.indexOf('function'), method.lastIndexOf('\n\t};') + 3).replace(/;$/, '');
    const sourceLogic = {name: 'source'};
    const sourceDrawing = {name: 'drawing', GetVisibleMMHeight: () => 200};
    const history = {depth: 0, TurnOff() { this.depth++; }, TurnOn() { this.depth--; }};
    const table = {m_bTurnOff: false};
    const context = {
        result: null,
        History: history,
        g_oTableId: table,
        g_dKoef_mm_to_pix: 4,
        editor: {WordControl: {m_oLogicDocument: sourceLogic, m_oDrawingDocument: sourceDrawing}},
        window: {NATIVE_EDITOR_ENJINE_SYNC_RECALC: 'previous'}
    };
    vm.runInNewContext('result=(' + expression + ')', context);
    for (const failure of [false, true]) {
        const originalDrawing = {name: 'merged-original'};
        const merged = {
            DrawingDocument: originalDrawing,
            RecalculateFromStart() {
                assert.equal(context.editor.WordControl.m_oLogicDocument, merged);
                assert.notEqual(this.DrawingDocument, originalDrawing);
                if (failure) throw Error('layout failed');
                this.DrawingDocument.OnStartRecalculate(1);
                this.DrawingDocument.OnRecalculatePage(0, {Width: 210, Height: 297});
                this.DrawingDocument.OnEndRecalculate(true, false);
            }
        };
        const receiver = {};
        if (failure) assert.throws(() => context.result.call(receiver, merged), /layout failed/);
        else {
            assert.equal(context.result.call(receiver, merged), true);
            assert.equal(receiver.printedDocument.DrawingDocument.m_arrPages[0].width_mm, 210);
            assert.equal(receiver.printedDocument.DrawingDocument.m_arrPages[0].height_mm, 297);
        }
        assert.equal(context.editor.WordControl.m_oLogicDocument, sourceLogic);
        assert.equal(context.editor.WordControl.m_oDrawingDocument, sourceDrawing);
        assert.equal(merged.DrawingDocument, originalDrawing);
        assert.equal(context.window.NATIVE_EDITOR_ENJINE_SYNC_RECALC, 'previous');
        assert.equal(history.depth, 0);
        assert.equal(table.m_bTurnOff, false);
    }
});
