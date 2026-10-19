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
const source = fs.readFileSync(path.join(__dirname, '../../word/Editor/Table.js'), 'utf8');
let geometryWritePosition = 0;
const geometryHistory = {Index: -1, Points: [], BinaryWriter: {GetCurPosition: () => geometryWritePosition}};
function method(name)
{
	const start = source.indexOf('CTable.prototype.' + name + ' = function(');
	const end = source.indexOf('\nCTable.prototype.', start + 1);
	assert(start > 0 && end > start);
	return vm.runInNewContext('(function(){let CTable = function(){};' + source.slice(start, end)
		+ '\n;return CTable.prototype.' + name + ';})()', {table_Selection_Text: 1, Asc: {linerule_Exact: 2}, AscCommon: {History: geometryHistory}, AscWord: {TableDirection: {reflectRulerMarkup() {}}}});
}
const hitRow = method('private_GetCellInRowByX');
const moveRTL = method('private_MoveCursorHorizontal');
function row(intervals)
{
	return {Index: 0, Get_CellsCount: () => intervals.length,
		Get_CellInfo: i => ({X_grid_start: intervals[i][0], X_grid_end: intervals[i][1]}),
		Get_Cell: i => ({Index: i})};
}
test('vertical navigation keeps the physical column in either logical order and at a nonzero page origin', () => {
	const table = {RowsInfo: [{StartPage: 0, Pages: 1}], Pages: [{X: 60}],
		CurCell: {Content: {GetRelativeStartPage: () => 0, GetRelativeCurrentPage: () => 0, GetAbsoluteCurrentPage: () => 0}}, GetAbsoluteStartPage: () => 0};
	const ltr = row([[0, 30], [30, 70], [70, 120]]);
	const rtl = row([[90, 120], [50, 90], [0, 50]]);
	for (const [x, expectedLTR, expectedRTL] of [[75, 0, 2], [110, 1, 1], [165, 2, 0]])
	{
		assert.equal(hitRow.call(table, ltr, x).Index, expectedLTR);
		assert.equal(hitRow.call(table, rtl, x).Index, expectedRTL);
	}
});
test('row navigation clamps to the target page and handles a merged interval and points outside the row', () => {
	const table = {RowsInfo: [{StartPage: 1, Pages: 1}], Pages: [{X: 10}, {X: 80}],
		CurCell: {Content: {GetRelativeStartPage: () => 0, GetRelativeCurrentPage: () => 0, GetAbsoluteCurrentPage: () => 0}}, GetAbsoluteStartPage: () => 0};
	const merged = row([[70, 120], [0, 70]]);
	for (const [x, expected] of [[20, 1], [100, 1], [180, 0], [240, 0]])
		assert.equal(hitRow.call(table, merged, x).Index, expected);
});
test('table routing preserves paragraph direction and normalizes legacy successful boundary returns', () => {
	for (const tableRTL of [false, true])
	for (const paragraphRTL of [false, true])
	for (const physicalRight of [false, true])
	{
		const logicalRight = physicalRight !== tableRTL;
		const characterRight = physicalRight !== paragraphRTL;
		const calls = [];
		let paragraphMoved = true;
		const content = {
			MoveCursorLeft: () => {calls.push('paragraph-left'); return paragraphMoved;},
			MoveCursorRight: () => {calls.push('paragraph-right'); return paragraphMoved;}
		};
		const table = {LogicDocument: {_tableCursorMove: {physicalRight, characterRight}},
			IsCellSelection: () => false, CurCell: {Index: 1, Row: {Index: 0}, Content: content},
			Selection: {StartPos: {}, EndPos: {}},
			MoveCursorLeft: (...args) => {calls.push(['table-previous', args]);},
			MoveCursorRight: (...args) => {calls.push(['table-next', args]);}};
		assert.equal(moveRTL.call(table, true, false, undefined, logicalRight), true);
		assert.equal(calls.length, 1);
		assert.equal(calls[0], characterRight ? 'paragraph-right' : 'paragraph-left');
		assert.equal(table.Selection.Type, 1);
		paragraphMoved = false;
		calls.length = 0;
		assert.equal(moveRTL.call(table, false, false, undefined, logicalRight), true);
		assert.equal(calls[1][0], logicalRight ? 'table-next' : 'table-previous');
		assert.deepEqual(calls[1][1].slice(-2), [true, true]);
	}
});
test('a single-cell selection collapses to the physical arrow edge even when the cell identity stays the same', () => {
	for (const physicalRight of [false, true])
	{
		let x;
		const content = {GetCurPosXY: () => ({X: 55, Y: 30}), RecalculateCurPos: () => ({X: 55, Y: 30}),
			GetRelativeStartPage: () => 0, GetRelativeCurrentPage: () => 0, GetAbsoluteCurrentPage: () => 0, GetAbsoluteStartPage: () => 0,
			MoveCursorToXY: value => {x = value;}};
		const table = {LogicDocument: {_tableCursorMove: {physicalRight, characterRight: physicalRight}},
			IsCellSelection: () => true, CurCell: {Content: content},
			MoveCursorLeft: () => true, MoveCursorRight: () => true};
		assert.equal(moveRTL.call(table, false, false, undefined, !physicalRight), true);
		assert.equal(x, physicalRight ? 10000 : -10000);
	}
});
test('actual ruler markup duplicates direction mapping and preserves its local axis through translations', () => {
	const metafile = fs.readFileSync(path.join(__dirname, '../../common/Drawings/Metafile.js'), 'utf8');
	const start = metafile.indexOf('function CTableMarkup(Table)');
	const end = metafile.indexOf('function CTableOutline(', start);
	const Markup = vm.runInNewContext('(function(){' + metafile.slice(start, end) + ';return CTableMarkup;})()');
	const original = new Markup({});
	original.X = 50; original.Cols = [30, 40]; original.Margins = [{Left: 2, Right: 3}];
	original.Rows = [{Y: 20, H: 10}]; original.TableDirectionRTL = true; original.TableDirectionAxis = 170; original.TableDirectionGeometry = 'snapshot';
	const copy = original.CreateDublicate();
	assert.equal(copy.TableDirectionGeometry, 'snapshot'); assert.equal(copy.TableDirectionRTL, true); assert.equal(copy.TableDirectionAxis, 170);
	copy.TransformX = 10; copy.TransformY = 5; copy.CorrectFrom();
	assert.equal(copy.X, 60); assert.equal(copy.Rows[0].Y, 25); assert.equal(copy.TableDirectionAxis, 170);
	copy.CorrectTo();
	assert.equal(copy.X, 50); assert.equal(copy.Rows[0].Y, 20); assert.equal(original.X, 50);
});
test('cell entry uses a freshly recalculated target line after a legacy undefined-success return', () => {
	let entry;
	const oldCell = {Content: {MoveCursorLeft: () => false}};
	const target = {Content: {GetCurPosXY: () => ({X: 1, Y: 0}),
		RecalculateCurPos: () => ({X: 20, Y: 99}),
		GetRelativeStartPage: () => 0, GetRelativeCurrentPage: () => 0, GetAbsoluteCurrentPage: () => 0, GetAbsoluteStartPage: () => 0,
		MoveCursorToXY: (x, y) => {entry = {x, y};}}};
	const table = {LogicDocument: {_tableCursorMove: {physicalRight: false, characterRight: false}},
		IsCellSelection: () => false, CurCell: oldCell,
		MoveCursorRight: () => {table.CurCell = target;}};
	assert.equal(moveRTL.call(table, false, false, undefined, true), true);
	assert.deepEqual(entry, {x: 10000, y: 99});
});

test('RTL resize uses continuous screen displacement across pages, scrolling and translated transforms', () => {
	const drawing = fs.readFileSync(path.join(__dirname, '../../word/Drawing/DrawingDocument.js'), 'utf8');
	const start = drawing.indexOf('this.getResizeDelta = function(');
	const end = drawing.indexOf('\n\t};', start) + 5;
	const getDelta = vm.runInNewContext('(function(){' + drawing.slice(start, end).replace('this.getResizeDelta', 'const getResizeDelta')
		+ ';return getResizeDelta;})()', {g_dKoef_pix_to_mm: 1,
		global_MatrixTransformer: {IsIdentity: () => false, Invert: () => ({
			TransformPointX: (x, y) => x / 2 + 40, TransformPointY: (x, y) => y / 2 + 80})},
		AscWord: {TableDirection: {resizePointerDelta: (a, b) => ({W: a.X - b.X, H: b.Y - a.Y})}}});
	const wordControl = {m_nZoomValue: 100};
	const tracker = {ResizeRTL: true, CurPos: {X: 100, Y: 200}};
	const delta = getDelta.call(tracker, 80, 230, wordControl);
	assert.equal(delta.W, 20); assert.equal(delta.H, 30);
	tracker.TableMatrix = {};
	const transformed = getDelta.call(tracker, 80, 230, wordControl);
	assert.equal(transformed.W, 10); assert.equal(transformed.H, 15);
	tracker.ResizeRTL = false; tracker.TrackOffsetX = 2; tracker.TrackOffsetY = 3;
	assert.equal(getDelta.call(tracker, 80, 230, wordControl).H, 27);
});
test('transformed resize paint extends southwest for RTL and southeast for LTR', () => {
	const drawing = fs.readFileSync(path.join(__dirname, '../../word/Drawing/DrawingDocument.js'), 'utf8');
	const start = drawing.indexOf('var _distX = (resizeRTL');
	const end = drawing.indexOf('\n\t\t\toverlay.CheckPoint', start);
	assert(start > 0 && end > start);
	const vertices = vm.runInNewContext('(function(resizeRTL, _ft, rPR){let _xLast = 30, _yLast = 70;'
		+ drawing.slice(start, end) + ';return _arr;})()'.replace('})()', '})'));
	// A translated scale stands in for the already composed page/object matrix.
	const transform = {GetScaleValue: () => 2, TransformPointX: (x, y) => x * 2 + 10,
		TransformPointY: (x, y) => y * 2 + 20};
	for (const rtl of [false, true])
	for (const dpr of [1, 2])
	{
		const points = vertices(rtl, transform, dpr);
		assert.equal(points[2] - points[0], rtl ? -6 * dpr : 6);
		assert.equal(points[5] - points[1], rtl ? 6 * dpr : 6);
		assert.equal(points[7] - points[1], rtl ? 6 * dpr : 6);
	}
});

test('pointer snapshots invalidate on source revisions while previews are neutral and validation stays constant-time', () => {
	const capture = method('GetInteractionGeometry'), current = method('IsInteractionGeometryCurrent');
	const table = {Distance: {L: 1, T: 2, R: 3, B: 4}, Inline: true, AllowOverlap: false, RecalcInfo: {InteractionRevision: 0}, TableGrid: [30, 40], PositionH: {Value: 10},
		PositionV: {Value: 20}, Pages: [{}], IsRTL: () => true};
	table.Parent = {Content: [table]};
	let matrix = {sx: 1, shx: 0, shy: 0, sy: 1, tx: 0, ty: 0}, documentRecalc = 1;
	table.Get_ParentTextTransform = () => ({...matrix});
	table.LogicDocument = {GetRecalcId: () => documentRecalc};
	const snapshot = capture.call(table);
	table.Selection = {Use: true}; table.Markup = {X: 88, Cols: [30, 33]};
	assert(current.call(table, snapshot));
	// Capture may find the parent index once. Mouse movement must not search it
	// again, inspect any rows/cells or serialize the table.
	table.Parent.Content.indexOf = () => {throw Error('hot-path parent search');};
	table.Content = new Proxy([], {get() {throw Error('hot-path table traversal');}});
	for (let i = 0; i < 1000; ++i) assert(current.call(table, snapshot));
	for (const mutate of [() => {table.Inline = !table.Inline;}, () => {table.AllowOverlap = !table.AllowOverlap;},
		() => {table.Distance.L++;}, () => {geometryWritePosition++;},
		() => {++table.RecalcInfo.InteractionRevision;},
		() => {table.PositionH.Value++;}, () => {documentRecalc++;},
		() => {matrix.tx++;}, () => {table.TableGrid = [30, 40];},
		() => {table.Pages = [{}];}, () => {table.TableStyle = 'changed';},
		() => {table.RecalcInfo = {InteractionRevision: table.RecalcInfo.InteractionRevision};}])
	{
		delete table.Parent.Content.indexOf;
		const before = capture.call(table); mutate(); assert.equal(current.call(table, before), false);
	}
	const before = capture.call(table); table.Parent.Content.unshift({});
	assert.equal(current.call(table, before), false);
});
test('real recalculation invalidators advance the geometry revision monotonically including Reset(false)', () => {
	const source = fs.readFileSync(path.join(__dirname, '../../word/Editor/Table/TableRecalculate.js'), 'utf8');
	const a = source.indexOf('function CTableRecalcInfo()'), b = source.indexOf('function CTableRecalculateObject()', a);
	const Info = vm.runInNewContext('(function(){' + source.slice(a, b) + ';return CTableRecalcInfo;})()');
	const info = new Info();
	for (const invalidate of [() => info.RecalcBorders(), () => info.Add_Cell({Get_Id: () => 'cell'}),
		() => info.Recalc_AllCells(), () => info.Reset(true), () => info.Reset(false)])
	{
		const before = info.InteractionRevision; invalidate(); assert(info.InteractionRevision > before);
	}
});

test('exact outline resize applies the delta once, including repeated header contributions', () => {
	const resize = method('private_ResizeExactRowsForOutline');
	function row(value) {return {Content: [], GetCellSpacing: () => null, GetHeight: () => ({Value: value, HRule: 2}), SetHeight: v => {value = v;}};}
	const table = {Content: [row(10), row(20)], GetMinRowHeight: () => 4.5, RowsInfo: [{H: [14], TopDy: [0]}, {H: [24], TopDy: [0]}], Pages: [{}, {}],
		HeaderInfo: {PageIndex: 0, Count: 1, Pages: [{Draw: false}, {Draw: true}]}};
	assert.equal(resize.call(table, 5.2), true);
	const a = table.Content[0].GetHeight().Value - 10, b = table.Content[1].GetHeight().Value - 20;
	assert(Math.abs(a - 1.4) < 1e-9); assert(Math.abs(b - 2.4) < 1e-9);
	assert(Math.abs(a * 2 + b - 5.2) < 1e-9);
	table.Content[1].GetHeight = () => ({Value: 20, HRule: 1});
	assert.equal(resize.call(table, 3), false);
});

test('a multi-element range collapses by document order without forcing the paragraph physical edge', () => {
	let args;
	const table = {Parent: {Selection: {Use: true, StartPos: 1, EndPos: 2}},
		LogicDocument: {_tableCursorMove: {physicalRight: false, characterRight: false}},
		MoveCursorLeft: (...a) => {args = a;}};
	assert.equal(moveRTL.call(table, false, false, undefined, false), true);
	assert.deepEqual(args, [false, false, true]);
});
test('exact row shrinking redistributes a constrained row instead of losing requested delta', () => {
	const resize = method('private_ResizeExactRowsForOutline');
	function row(value) {return {Content: [], GetCellSpacing: () => null,
		GetHeight: () => ({Value: value, HRule: 2}), SetHeight: v => {value = v;}};}
	const table = {Content: [row(5), row(100)], RowsInfo: [{H: [15], TopDy: [0]}, {H: [100], TopDy: [0]}],
		Pages: [{}], GetMinRowHeight: () => 4.5, HeaderInfo: {PageIndex: 0, Count: 0, Pages: []}};
	assert.equal(resize.call(table, -80), true);
	assert.equal(table.Content[0].GetHeight().Value, 4.5);
	assert.equal(table.Content[1].GetHeight().Value, 20.5);
});

test('starting text selection preserves the current caret instead of selecting the cell first', () => {
	const start = method('StartSelectionFromCurPos');
	let caret = 7, recursiveStart;
	const table = {IsRTL: () => true, Selection: {StartPos: {}, EndPos: {}}, CurCell: {Index: 1, Row: {Index: 2},
		Content: {GetRelativeStartPage: () => 1, GetRelativeCurrentPage: () => 1, GetAbsoluteCurrentPage: () => 5, StartSelectionFromCurPos: () => {recursiveStart = caret;}}},
		LogicDocument: {GetCursorRealPosition: () => ({X: 30, Y: 40})},
		GetAbsoluteStartPage: () => 5, private_UpdateSelectedCellsArray: () => {caret = 99;},
		private_SetSelectionData: data => {assert.equal(data, null);}};
	start.call(table);
	assert.equal(recursiveStart, 7); assert.equal(caret, 7);
	assert.equal(table.Selection.Type, 1); assert.equal(table.Selection.StartPos.PageIndex, 2); assert.equal(table.Selection.EndPos.PageIndex, 2);
	assert.equal(table.Selection.StartPos.X, 30); assert.equal(table.Selection.StartPos.Y, 40);
});

test('vertical row origin uses the local column fragment even on the same absolute page', () => {
	const table = {RowsInfo: [{StartPage: 0, Pages: 2}], Pages: [{X: 10}, {X: 80}],
		CurCell: {Content: {GetRelativeStartPage: () => 1, GetRelativeCurrentPage: () => 0,
			GetAbsoluteCurrentPage: () => 0}}, GetAbsoluteStartPage: () => 0};
	const target = row([[0, 30], [30, 70]]);
	assert.equal(hitRow.call(table, target, 100).Index, 0);
});

test('selection bounds map local column fragments to physical pages and distinct column origins', () => {
	const bounds = method('GetSelectionBounds');
	const anchor = method('GetSelectionAnchorPos');
	for (const rtl of [false, true])
	{
		const metrics = rtl ? {X_cell_start: 50, X_cell_end: 70} : {X_cell_start: 10, X_cell_end: 30};
		const cells = [{Row: 0, Cell: 0}, {Row: 1, Cell: 0}];
		const table = {IsRTL: () => rtl, IsCellSelection: () => true, GetSelectionArray: () => cells,
			Content: [0, 1].map(() => ({Get_Cell: () => ({Metrics: metrics})})),
			RowsInfo: [{StartPage: 0, Pages: 1, X0: 2, Y: [20], H: [10]},
				{StartPage: 1, Pages: 1, X0: 4, Y: [undefined, 30], H: [undefined, 12]}],
			Pages: [{X: 20}, {X: 110}], GetAbsoluteStartPage: () => 3,
			GetAbsolutePage: local => local < 2 ? 3 : 4,
			Selection: {StartPos: {Pos: cells[1]}, EndPos: {Pos: cells[0]}}};
		const result = bounds.call(table);
		assert.equal(result.Start.Page, 3); assert.equal(result.End.Page, rtl ? 3 : 4);
		assert.equal(result.Start.X, 20 + (rtl ? 0 : 2) + metrics.X_cell_start);
		assert.equal(result.End.X, (rtl ? 110 : 22) + metrics.X_cell_start);
		assert.equal(result.Direction, -1);
		cells.splice(0, 1);
		const position = anchor.call(table);
		assert.equal(position.Page, rtl ? 3 : 4); assert.equal(position.X0, (rtl ? 110 : 0) + metrics.X_cell_start);
		assert.equal(position.X1, (rtl ? 110 : 0) + metrics.X_cell_end);
	}
});


test('actual ruler commit validates the captured revision without recapturing geometry', () => {
 const rulers=fs.readFileSync(path.join(__dirname,'../../word/Drawing/Rulers.js'),'utf8');
 const start=rulers.indexOf('    this.SetTableProperties = function()'),end=rulers.indexOf('    this.SetColumnsProperties = function()',start);
 assert(start>0&&end>start);
 const guard=rulers.slice(start,end);
 const commit=vm.runInNewContext('(function(){'+guard+';return this.SetTableProperties;}).call({})',
  {AscCommon:{changestype_Table_Properties:1},AscDFH:{historydescription_Document_SetTableMarkup_Hor:2}});
 for(const valid of [true,false]){
  let checks=0,actions=0,updates=0,applies=0;
  const snapshot={},markup={Table:{IsInteractionGeometryCurrent(value){checks++;assert.equal(value,snapshot);return valid;},GetInteractionGeometry(){throw Error('Hot path recapture');},Update_TableMarkupFromRuler(){applies++;}},CorrectTo(){},CorrectFrom(){}};
  const ruler={TableDirectionDrag:{markup,geometry:snapshot},m_oTableMarkup:markup,m_oWordControl:{m_oLogicDocument:{
   UpdateRulers(){updates++;},Document_Is_SelectionLocked:()=>false,StartAction(){actions++;},FinalizeAction(){},UpdateSelection(){},UpdateInterface(){}
  }}};
  commit.call(ruler);assert.equal(checks,1);assert.equal(actions,valid?1:0);assert.equal(applies,valid?1:0);assert.equal(updates,1);assert.equal(ruler.TableDirectionDrag,null);
 }
});


test('real ruler-to-table revalidation accepts its empty action and rejects an external recorded edit', () => {
 const rulers=fs.readFileSync(path.join(__dirname,'../../word/Drawing/Rulers.js'),'utf8');
 const start=rulers.indexOf('    this.SetTableProperties = function()'),end=rulers.indexOf('    this.SetColumnsProperties = function()',start);
 const commit=vm.runInNewContext('(function(){'+rulers.slice(start,end)+';return this.SetTableProperties;}).call({})',
  {AscCommon:{changestype_Table_Properties:1},AscDFH:{historydescription_Document_SetTableMarkup_Hor:2}});
 const capture=method('GetInteractionGeometry'),current=method('IsInteractionGeometryCurrent'),apply=method('Update_TableMarkupFromRuler');
 for(const external of [false,true])for(const replaceEmpty of [false,true]){
  geometryHistory.Index=0;geometryHistory.Points=[{Items:replaceEmpty?[]:[{}]}];
  let entersTable=0,actions=0,updates=0;
  const table={Inline:true,AllowOverlap:false,Distance:{L:1,T:2,R:3,B:4},RecalcInfo:{InteractionRevision:0},TableGrid:[30,40],
   PositionH:{Value:1},PositionV:{Value:2},Pages:[{X:0}],IsRTL:()=>true,Get_ParentTextTransform:()=>null,
   LogicDocument:{GetRecalcId:()=>1},Get_CompiledPr(){entersTable++;return{TablePr:{}};},
   Content:[{Get_CellInfo:()=>({StartGridCol:1})}],Markup:{X:0,Cols:[30,40]}};
  table.Parent={Content:[table]};table.IsInteractionGeometryCurrent=snapshot=>current.call(table,snapshot);
  table.Update_TableMarkupFromRuler=(...args)=>apply.apply(table,args);
  const snapshot=capture.call(table),markup={Table:table,TableDirectionRTL:true,TableDirectionGeometry:snapshot,TableDirectionAxis:70,
   Internal:{PageNum:0,RowIndex:0},X:0,Cols:[30,40],CorrectTo(){},CorrectFrom(){},CreateDublicate(){return this;}};
  if(external){geometryHistory.Points[0].Items.push({});geometryWritePosition++;}
  const ruler={TableDirectionDrag:{markup,geometry:snapshot},m_oTableMarkup:markup,DragTablePos:1,m_oWordControl:{m_oLogicDocument:{
   UpdateRulers(){updates++;},Document_Is_SelectionLocked:()=>false,StartAction(){actions++;if(replaceEmpty)geometryHistory.Points[0]={Items:[]};else{geometryHistory.Points.push({Items:[]});geometryHistory.Index++;}},FinalizeAction(){},UpdateSelection(){},UpdateInterface(){}
  }}};
  commit.call(ruler);assert.equal(actions,external?0:1);assert.equal(entersTable,external?0:1);assert.equal(updates,1);
 }
});

test('LTR presentation selection keeps the original path without a Word table-stack capability', () => {
 const start=method('StartSelectionFromCurPos');let original=0,recursive=0;
 const table={IsRTL:()=>false,Selection:{StartPos:{},EndPos:{}},CurCell:{Index:1,Row:{Index:0},Content:{StartSelectionFromCurPos(){recursive++;}}},
  LogicDocument:{GetCursorRealPosition:()=>({X:10,Y:20})},
  private_UpdateSelectedCellsArray(){original++;},private_SetSelectionData(){throw Error('Non-Word LTR must keep the original selection path');}};
 start.call(table);assert.equal(original,1);assert.equal(recursive,1);assert.equal(table.Selection.Type,1);assert.equal(table.Selection.Use,true);
});
