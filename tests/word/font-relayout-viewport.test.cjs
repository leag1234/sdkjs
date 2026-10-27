/**
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const root = path.join(__dirname, '../..');
const documentSource = fs.readFileSync(path.join(root, 'word/Editor/Document.js'), 'utf8');
const drawingSource = fs.readFileSync(path.join(root, 'word/Drawing/DrawingDocument.js'), 'utf8');
const htmlSource = fs.readFileSync(path.join(root, 'word/Drawing/HtmlPage.js'), 'utf8');

// Run the actual document save/load/anchor/recalculate and wheel implementations.
// Geometry, pagination scheduling and the drawing sink are fixture boundaries;
// no replacement font-refresh or viewport-preservation algorithm is used here.
function fixture({caret = 0, range = null, asynchronous = false, shift = 50, viewport = 887, pageHeight = Infinity} = {})
{
	const f = {viewport, visibleViewport: viewport, caret, range, shift: 0, queue: [], locks: [], scrolls: [], cacheResets: 0,
		history: {points: ['existing-edit'], index: 0}, dirty: false, mutationCalls: [], saves: 0, restores: 0, recalculations: 0};
	const history = {
		Create_NewPoint: () => { f.mutationCalls.push('history point'); f.history.points.push('font refresh'); },
		Add: () => { f.mutationCalls.push('history change'); f.history.index++; }
	};
	const context = vm.createContext({
		CDocument: function() {},
		AscWord: {CRun: function() {}},
		AscCommon: {History: history, CollaborativeEditing: {Get_GlobalLockSelection: () => false}, g_clipboardBase: {IsWorking: () => false}, g_fontManager: {ClearRasterMemory() {}}, check_KeyboardEvent: e => { context.global_keyboardEvent = {KeyCode: e.keyCode}; },
			checkMouseWhell: e => ({x: 0, y: e.deltaY})},
		Asc: {editor: {isPdfEditor: () => false}},
		selectionflag_Numbering: 1,
		keydownresult_PreventKeyPress: 1, keydownresult_PreventDefault: 2, g_dKoef_mm_to_pix: 1,
		window: {},
		global_mouseEvent: {X: 0, Y: 0, AltKey: false, ShiftKey: false, CtrlKey: false, Sender: null}
	});
	const enumStart = documentSource.indexOf('AscWord.ViewPositionType = {');
	assert(enumStart >= 0);
	vm.runInContext(documentSource.slice(enumStart, documentSource.indexOf('};', enumStart) + 2), context);
	const names = ['RecalculateFromStart', 'UpdateCursorOnRecalculate', '_isSelectionVisible', 'Document_UpdateSelectionState',
		'UpdateSelection', 'private_UpdateSelection', 'controller_UpdateSelectionState', 'RecalculateCurPos', 'Set_TargetPos',
		'Save_DocumentStateBeforeLoadChanges', 'private_StoreViewPositions',
		'Load_DocumentStateAfterLoadChanges', 'CheckViewPosition'];
	for (const match of documentSource.matchAll(/CDocument\.prototype\.([A-Za-z0-9_]*LocalFont[A-Za-z0-9_]*)\s*=/g))
		if (!names.includes(match[1])) names.push(match[1]);
	for (const name of names)
	{
		const start = documentSource.indexOf('CDocument.prototype.' + name + ' =');
		assert(start >= 0, name);
		const end = documentSource.indexOf('\nCDocument.prototype.', start + 1);
		vm.runInContext(documentSource.slice(start, end < 0 ? undefined : end), context);
	}
	const doc = new context.CDocument();
	f.doc = doc;
	doc.CurPos = {X: 0, Y: 0, RealX: 0, RealY: 0, Type: 0};
	doc.Selection = {Start: false, Use: !!range, Flag: 0, UpdateOnRecalc: false,
		DragDrop: {Flag: false, Data: null}};
	doc.ViewPosition = null;
	doc.Pages = [];
	doc.FullRecalc = {Id: null, StartIndex: 0, ScrollToTarget: true};
	doc.History = history;
	doc.Create_NewHistoryPoint = history.Create_NewPoint;
	const pageOf = y => Number.isFinite(pageHeight) ? Math.floor(y / pageHeight) : 0;
	const localY = y => Number.isFinite(pageHeight) ? y % pageHeight : y;
	const absoluteY = (page, y) => Number.isFinite(pageHeight) ? page * pageHeight + y : y;
	const clone = value => value === null ? null : JSON.parse(JSON.stringify(value));
	function position(y)
	{
		const run = new context.AscWord.CRun();
		run.logicalY = y;
		run.State = {ContentPos: 0};
		const paragraph = {
			SaveSelectionState: () => ({caret: f.caret, range: clone(f.range)}),
			RemoveSelection: () => { f.range = null; },
			LoadSelectionState: state => { f.caret = state.caret; f.range = clone(state.range); },
			RecalculateCurPos: () => ({PageNum: pageOf(run.logicalY + f.shift), X: 0, Y: localY(run.logicalY + f.shift), Height: 12})
		};
		run.GetParagraph = () => paragraph;
		run.SetThisElementCurrentInParagraph = () => { f.caret = run.logicalY; };
		return [{Class: doc, Position: Math.max(0, Math.floor(y / 20))}, {Class: run, Position: 0}];
	}
	doc.Controller = {
		SaveDocumentStateBeforeLoadChanges: state => {
			f.saves++;
			state.testCaret = f.caret; state.testRange = clone(f.range);
			state.Pos = position(f.caret);
			state.StartPos = position(f.range ? f.range.start : f.caret);
			state.EndPos = position(f.range ? f.range.end : f.caret);
		},
		RestoreDocumentStateAfterLoadChanges: state => {
			f.restores++; f.caret = state.testCaret; f.range = clone(state.testRange);
		}
	};
	doc.CollaborativeEditing = {WatchDocumentPositionsByState() {}, UpdateDocumentPositionsByState() {}};
	doc.GetSelectedElementsInfo = () => ({GetSingleCell: () => null});
	doc.RemoveSelection = () => { doc.Selection.Use = false; f.range = null; };
	doc.SetDocPosType = type => { doc.CurPos.Type = type; };
	doc.TargetPos = {};
	doc.Action = {Start: false, UndoRedo: false, UpdateStates: false, ScrollToTarget: false};
	doc.RemoveEmptySelection = true;
	doc.HandleOformSelectionInEditMode = () => false;
	doc.CheckFixedFormSelectionInEditMode = () => {};
	doc.IsFillingFormMode = doc.IsMovingTableBorder = () => false;
	doc.UpdateDocumentOutlinePosition = doc.Document_UpdateCopyCutState = () => {};
	doc.private_UpdateTracks = doc.private_CheckCurPage = doc.Document_UpdateRulersState = () => {};
	f.selectionUpdates = 0;
	doc.Controller.UpdateSelectionState = () => {
		f.selectionUpdates++;
		doc.controller_UpdateSelectionState();
	};
	doc.Controller.RecalculateCurPos = () => doc.DrawingDocument.scrollToTarget();
	doc.IsSelectionUse = doc.IsTextSelectionUse = () => doc.Selection.Use;
	doc.IsSelectionEmpty = () => !f.range || f.range.start === f.range.end;
	doc.GetSelectionBounds = () => ({
		Direction: 1,
		Start: {Page: 0, Y: f.range ? f.range.start : f.caret, H: 12},
		End: {Page: 0, Y: f.range ? f.range.end : f.caret, H: 12}
	});
	doc.GetDocumentPositionByXY = (page, _x, y) => position(Math.ceil((absoluteY(page, y) - f.shift) / 20) * 20);
	doc.private_GetXYByDocumentPosition = pos => ({Page: pageOf(pos[1].Class.logicalY + f.shift), X: 0, Y: localY(pos[1].Class.logicalY + f.shift), H: 12});
	const control = {
		m_oApi: {bInit_word_control: true, checkLastWork() {}, isLongAction: () => false},
		StartUpdateOverlay() {}, EndUpdateOverlay() {}, paintMessageLoop: {requestUrgent() {}},
		m_bIsHorScrollVisible: false, m_bIsRePaintOnScroll: true,
		onMouseMove() {},
		ScrollToAbsolutePosition: (_x, y, page) => { f.viewport = absoluteY(page, y); f.scrolls.push(['anchor', f.viewport]); }
	};
	control.m_oLogicDocument = doc;
	control.m_dScrollY = f.viewport;
	control.m_dScrollX_max = control.m_dScrollY_max = 10000;
	control.m_oEditor = {AbsolutePosition: {L: 0, T: 0, R: 800, B: 600}};
	control.CheckZoom = control.CalculateDocumentSize = control.OnCalculatePagesPlace = () => {};
	control.OnScroll = () => { f.visibleViewport = control.m_dScrollY; };
	control.m_oApi.sync_zoomChangeCallback = () => {};

	control.m_oScrollVerApi = {scrollBy: (_x, y) => { f.viewport += y; f.verticalScroll({getIsLockedMouse: () => false}, f.viewport, 10000); }, getCurScrolledY: () => f.viewport,
		scrollToY: y => { f.viewport = y; }};
	control.m_oScrollHorApi = {scrollBy() {}, getCurScrolledX: () => 0, scrollToX() {}};
	doc.DrawingDocument = {
		m_oWordControl: control, m_lCurrentPage: -1,
		UpdateTargetTransform() {}, TargetEnd() {}, TargetStart() {}, TargetShow() {},
		SelectEnabled() {}, SelectShow() {},
		ConvertCoordsFromCursor2: () => ({Page: 0, X: 0, Y: f.viewport}),
		ConvertCoordsToCursor: () => ({X: 400, Y: 300}),
		GetVisibleRegion: () => [{Page: pageOf(f.visibleViewport), Y: localY(f.visibleViewport)},
			{Page: pageOf(f.visibleViewport + 300), Y: localY(f.visibleViewport + 300)}],
		scrollToTarget: () => { if (!f.locked) { f.viewport = f.caret; f.scrolls.push(['caret', f.caret]); } }
	};
	control.m_oDrawingDocument = doc.DrawingDocument;
	Object.defineProperty(control.m_oApi, 'isLockScrollToTarget', {get: () => doc.Api.isLockScrollToTarget});
	const targetStart = drawingSource.indexOf('this.UpdateTarget = function (x, y, pageIndex)');
	const targetNext = drawingSource.slice(targetStart + 1).search(/\n\tthis\.[A-Za-z0-9_]+\s*=/);
	assert(targetStart >= 0 && targetNext >= 0);
	context.drawing = doc.DrawingDocument;
	f.deferredTarget = vm.runInContext('(function(){' + drawingSource.slice(targetStart, targetStart + 1 + targetNext) +
		'\n;return this.UpdateTarget;}).call(drawing)', context).bind(doc.DrawingDocument);

	doc.Api = {isLockScrollToTarget: false, setDocumentModified: value => { f.mutationCalls.push('dirty'); f.dirty = value; }, asc_LockScrollToTarget: value => { doc.Api.isLockScrollToTarget = value; f.locked = value; f.locks.push(value); }};
	doc.private_UpdateCursorXY = () => {};
	doc.Reset_RecalculateCache = () => { f.cacheResets++; };
	doc.Document_UpdateInterfaceState = () => {};
	doc.RecalculateWithParams = () => {
		const generation = ++f.recalculations;
		doc.FullRecalc.Id = 1; doc.FullRecalc.ScrollToTarget = true;
		const finish = () => {
			if (generation !== f.recalculations) return; // Actual SDK discards superseded timer generations.
			f.shift += shift;
			doc.UpdateCursorOnRecalculate();
			doc.FullRecalc.Id = null;
			doc.CheckViewPosition();
		};
		if (asynchronous) f.queue.push(finish); else finish();
	};
	context.control = control;
	context.oThis = control;
	const wheelStart = htmlSource.indexOf('this.onMouseWhell = function(e)');
	const wheelEnd = htmlSource.indexOf('this.SelectWheel =', wheelStart);
	assert(wheelStart >= 0 && wheelEnd > wheelStart);
	f.wheel = vm.runInContext('(function(){' + htmlSource.slice(wheelStart, wheelEnd) +
		';return this.onMouseWhell;}).call(control)', context);
	f.finish = () => { while (f.queue.length) f.queue.shift()(); };
	f.refresh = () => doc.RecalculateFromStart(true, true);
	const keyStart = htmlSource.indexOf('this.onKeyDown = function(e)');
	const keyEnd = htmlSource.indexOf('this.onKeyDownNoActiveControl =', keyStart);
	assert(keyStart >= 0 && keyEnd > keyStart);
	f.key = vm.runInContext('(function(){' + htmlSource.slice(keyStart, keyEnd) +
		';return this.onKeyDown;}).call(control)', context);
	doc.OnKeyDown = () => { f.caret = 1100; doc.DrawingDocument.scrollToTarget(); return 0; };
	for (const name of ['verticalScroll', 'zoom_Fire', 'onKeyPress'])
	{
		const start = htmlSource.indexOf('this.' + name + ' = function');
		const next = htmlSource.slice(start + 1).search(/\n\tthis\.[A-Za-z0-9_]+\s*=/);
		assert(start >= 0 && next >= 0, name);
		f[name] = vm.runInContext('(function(){' + htmlSource.slice(start, start + 1 + next) +
			'\n;return this.' + name + ';}).call(control)', context).bind(control);
	}
	f.scrollbar = y => { f.viewport = y; f.verticalScroll({getIsLockedMouse: () => true}, y, 10000); };
	f.before = {caret: f.caret, range: clone(f.range), history: clone(f.history), dirty: f.dirty};
	return f;
}

test('late font relayout retains the reading anchor with the untouched initial caret', () => {
	const f = fixture();
	f.refresh();
	assert.equal(f.viewport, 937);
	assert.equal(f.caret, 0);
	assert.equal(f.recalculations, 1);
	assert.equal(f.cacheResets, 1);
});
test('an offscreen clicked caret does not replace the reading viewport', () => {
	const f = fixture({caret: 320});
	f.refresh();
	assert.equal(f.viewport, 937);
	assert.equal(f.caret, 320);
});
test('noncollapsed selection survives the real state/anchor save and restore', () => {
	const f = fixture({caret: 320, range: {start: 280, end: 320}});
	f.refresh();
	assert.equal(f.viewport, 937);
	assert.deepEqual(f.range, f.before.range);
	assert.equal(f.doc.Selection.Use, true);
});
test('asynchronous pagination retains the anchor until its layout is available', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	assert.equal(f.viewport, 887, 'Starting the background calculation must not reveal the old caret');
	assert(f.doc.ViewPosition, 'SDK anchor must survive the asynchronous boundary');
	const anchor = f.doc.ViewPosition;
	f.doc.CheckViewPosition();
	assert.equal(f.doc.ViewPosition, anchor, 'An unfinished page cannot consume its anchor');
	assert.equal(f.scrolls.length, 0);
	assert.equal(f.viewport, 887);
	f.finish();
	assert.equal(f.viewport, 937);
	assert.equal(f.doc.ViewPosition, null);
});
test('a newer real wheel event during pagination replaces the stale reading anchor', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	f.wheel({deltaY: 400, ctrlKey: false, preventDefault() {}});
	f.finish();
	assert.equal(f.viewport, 1337, 'Keep the newer logical reading position after the font metrics shift');
});
test('newer caret and selection input during pagination is never restored to the old state', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	f.caret = 1100; f.range = {start: 1080, end: 1100}; f.doc.Selection.Use = true;
	f.finish();
	assert.equal(f.viewport, 937);
	assert.equal(f.caret, 1100);
	assert.deepEqual(f.range, {start: 1080, end: 1100});
});
test('background font relayout creates no dirty state or undo history entry', () => {
	const f = fixture({asynchronous: true, range: {start: 280, end: 320}});
	f.refresh(); f.finish();
	assert.deepEqual(f.history, f.before.history);
	assert.deepEqual(f.mutationCalls, [], 'Actual SDK/API mutation boundaries must not be called');
	assert.equal(f.dirty, f.before.dirty);
	assert.deepEqual(f.range, f.before.range);
	assert.equal(f.recalculations, 1);
});
test('ordinary explicit recalculation still follows its caret after background restoration ends', () => {
	const f = fixture();
	f.refresh();
	f.caret = 200;
	f.doc.RecalculateFromStart(true);
	assert.equal(f.viewport, 200);
});

test('a viewport crossing a page boundary follows the same logical anchor on the next page', () => {
	const f = fixture({asynchronous: true, viewport: 2970, pageHeight: 1000});
	f.refresh();
	f.doc.CheckViewPosition();
	assert.equal(f.viewport, 2970);
	f.finish();
	assert.equal(f.viewport, 3020);
	assert.equal(f.scrolls.length, 1);
});
test('font refresh does not take ownership of an existing foreign restore anchor', () => {
	const f = fixture({asynchronous: true});
	f.doc.Load_DocumentStateAfterLoadChanges(f.doc.Save_DocumentStateBeforeLoadChanges(false, true), false);
	const foreign = f.doc.ViewPosition;
	f.refresh();
	assert.equal(f.doc.ViewPosition, foreign);
	assert(!f.doc.LocalFontRelayoutViewPosition);
	f.wheel({deltaY: 400, ctrlKey: false, preventDefault() {}});
	assert.equal(f.doc.ViewPosition, foreign);
});
test('a refused or zero-delta wheel cannot recapture an anchor from partially changed geometry', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	const anchor = f.doc.ViewPosition;
	f.shift = 20;
	f.wheel({deltaY: 0, ctrlKey: false, preventDefault() {}});
	assert.equal(f.doc.ViewPosition, anchor);
	f.wheel({deltaY: 400, ctrlKey: true, preventDefault() {}});
	assert.equal(f.doc.ViewPosition, anchor);
});
test('a real keyboard navigation event takes priority over a pending font anchor', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	f.key({keyCode: 34, preventDefault() {}});
	f.finish();
	assert.equal(f.caret, 1100);
	assert.equal(f.viewport, 1100);
	assert.equal(f.doc.ViewPosition, null);
});
test('two overlapping font requests keep the pending logical anchor', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	const first = f.doc.ViewPosition;
	f.refresh();
	assert.deepEqual(f.doc.ViewPosition.AnchorPos, first.AnchorPos);
	assert.equal(f.doc.ViewPosition.Distance, first.Distance);
	assert.equal(f.doc.ViewPosition, f.doc.LocalFontRelayoutViewPosition);
	f.finish();
	assert.equal(f.viewport, 937);
	assert.equal(f.doc.ViewPosition, null);
	assert.equal(f.doc.LocalFontRelayoutViewPosition, null);
	assert.equal(f.scrolls.length, 1);
});
test('a font relayout preserves an already dirty history state', () => {
	const f = fixture({asynchronous: true});
	f.dirty = true;
	f.refresh(); f.finish();
	assert.equal(f.dirty, true);
	assert.deepEqual(f.history, f.before.history);
	assert.deepEqual(f.mutationCalls, []);
});

test('mouse-held scrollbar navigation recaptures the newer viewport', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	f.scrollbar(1400);
	f.finish();
	assert.equal(f.viewport, 1450);
});
test('programmatic pagination scroll clamps do not recapture the initial logical anchor', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	const anchor = f.doc.ViewPosition;
	f.verticalScroll({getIsLockedMouse: () => false}, 880, 10000);
	assert.equal(f.doc.ViewPosition, anchor);
	f.finish();
	assert.equal(f.viewport, 937);
});
test('ordinary zoom executes the real control method without relying on a scrollbar sender', () => {
	const f = fixture();
	assert.doesNotThrow(() => f.zoom_Fire(1, 100));
});

test('newer wheel input cannot resurrect the stale anchor when temporary page geometry is unavailable', () => {
	const f = fixture({asynchronous: true});
	f.refresh();
	f.doc.DrawingDocument.GetVisibleRegion = () => null;
	f.wheel({deltaY: 400, ctrlKey: false, preventDefault() {}});
	f.doc.FullRecalc.Id = null;
	f.doc.UpdateCursorOnRecalculate();
	assert.equal(f.viewport, 1287, 'The final cursor update cannot reveal the stale caret');
	f.finish();
	assert.equal(f.viewport, 1287);
	assert.equal(f.doc.ViewPosition, null);
	assert.equal(f.doc.LocalFontRelayoutViewPosition, null);
});

test('shared control wheel, scrollbar and keypress still work with a non-Word logic document', () => {
	const f = fixture();
	const control = f.doc.DrawingDocument.m_oWordControl;
	control.m_oLogicDocument = {OnKeyPress: () => false};
	assert.doesNotThrow(() => f.wheel({deltaY: 100, ctrlKey: false, preventDefault() {}}));
	assert.doesNotThrow(() => f.scrollbar(1200));
	assert.doesNotThrow(() => f.onKeyPress({keyCode: 65, preventDefault() {}}));
});

test('font relayout still updates the real selection UI without revealing its offscreen caret', () => {
	const f = fixture({caret: 320});
	f.refresh();
	assert.equal(f.selectionUpdates, 1);
	assert.equal(f.viewport, 937);
	assert.equal(f.doc.Api.isLockScrollToTarget, false);
});
test('a failed font selection UI update releases its temporary scroll lock', () => {
	const f = fixture();
	f.doc.Controller.UpdateSelectionState = () => { throw Error('Synthetic selection UI failure'); };
	assert.throws(() => f.refresh(), /Synthetic selection UI failure/);
	assert.equal(f.doc.Api.isLockScrollToTarget, false);
});

test('the actual deferred drawing target retains the font UI scroll lock until paint and releases on new input', () => {
	const f = fixture({caret: 320});
	const drawing = f.doc.DrawingDocument;
	drawing.UpdateTargetFromPaint = false;
	drawing.m_lCurrentPage = 0;
	f.doc.Controller.RecalculateCurPos = () => f.deferredTarget(0, f.caret, 0);
	f.refresh();
	assert.equal(f.viewport, 937);
	assert.equal(drawing.isDisableScrollToTarget, true);
	assert.equal(f.doc.Api.isLockScrollToTarget, false);
	f.deferredTarget(0, 1100, 0);
	assert.equal(drawing.isDisableScrollToTarget, false, 'A later ordinary input must re-enable target scrolling');
});

test('the font UI phase preserves an already held caller scroll lock', () => {
	const f = fixture({asynchronous: true});
	f.doc.Api.asc_LockScrollToTarget(true);
	f.refresh();
	assert.equal(f.doc.Api.isLockScrollToTarget, true);
	f.doc.Api.asc_LockScrollToTarget(false);
	f.finish();
	assert.equal(f.viewport, 937);
});
