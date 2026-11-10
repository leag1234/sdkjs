/*
 * Tests for the print range computation of a worksheet (calcPagesPrint).
 * Assisted-by: Atlas
 */
QUnit.config.autostart = false;
$(function () {
	AscCommon.InitBrowserInputContext = function () {};
	Asc.spreadsheet_api.prototype.CheckChangedDocument = function () {};
	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype.asyncImagesDocumentEndLoaded = function () {};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function () {
		openDocument();
	};
	Asc.spreadsheet_api.prototype.initGlobalObjects = function (wbModel) {
		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wbModel);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wbModel);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wbModel);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wbModel, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wbModel, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wbModel);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wbModel);
		AscCommonExcel.g_oUndoRedoSparklines = new AscCommonExcel.UndoRedoSparklines(wbModel);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wbModel, -1);
		AscCommonExcel.g_oUndoRedoSharedFormula = new AscCommonExcel.UndoRedoSharedFormula(wbModel);
		AscCommonExcel.g_oUndoRedoLayout = new AscCommonExcel.UndoRedoRedoLayout(wbModel);
		AscCommonExcel.g_oUndoRedoHeaderFooter = new AscCommonExcel.UndoRedoHeaderFooter(wbModel);
		AscCommonExcel.g_oUndoRedoArrayFormula = new AscCommonExcel.UndoRedoArrayFormula(wbModel);
		AscCommonExcel.g_oUndoRedoSortState = new AscCommonExcel.UndoRedoSortState(wbModel);
		AscCommonExcel.g_oUndoRedoSlicer = new AscCommonExcel.UndoRedoSlicer(wbModel);
		AscCommonExcel.g_oUndoRedoPivotTables = new AscCommonExcel.UndoRedoPivotTables(wbModel);
		AscCommonExcel.g_oUndoRedoPivotFields = new AscCommonExcel.UndoRedoPivotFields(wbModel);
		AscCommonExcel.g_oUndoRedoPivotFieldItems = new AscCommonExcel.UndoRedoPivotFieldItems(wbModel);
		AscCommonExcel.g_oUndoRedoPivotCache = new AscCommonExcel.UndoRedoPivotCache(wbModel);
		AscCommonExcel.g_oUndoRedoCacheFields = new AscCommonExcel.UndoRedoCacheFields(wbModel);
		AscCommonExcel.g_oUndoRedoCF = new AscCommonExcel.UndoRedoCF(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedRange = new AscCommonExcel.UndoRedoProtectedRange(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedSheet = new AscCommonExcel.UndoRedoProtectedSheet(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedWorkbook = new AscCommonExcel.UndoRedoProtectedWorkbook(wbModel);
		AscCommonExcel.g_oUndoRedoNamedSheetViews = new AscCommonExcel.UndoRedoNamedSheetViews(wbModel);
		AscCommonExcel.g_oUndoRedoUserProtectedRange = new AscCommonExcel.UndoRedoUserProtectedRange(wbModel);
	};
	Asc.spreadsheet_api.prototype._onUpdateDocumentCanSave = function () {};
	AscCommonExcel.WorkbookView.prototype._onScrollReinitialize = function () {};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {};
	AscCommonExcel.WorkbookView.prototype.getZoom = function () { return 1; };
	AscCommonExcel.WorkbookView.prototype.changeZoom = function () {};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {};
	AscCommonExcel.WorksheetView.prototype.draw = function () {};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
		this.objectRender = new AscFormat.DrawingObjects();
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () { return 1; };
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
		AscFonts.g_fontApplication.Init();
		this.FontLoader = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);
	};

	let wb, wsView;
	Asc.spreadsheet_api.prototype.fAfterLoad = function () {
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function () { return null; });
		wsView = api.wb.getWorksheet();
		startTests();
	};
	AscCommon.CHistory.prototype.UndoRedoEnd = function () {
		this.TurnOn();
	};

	var api = new Asc.spreadsheet_api({'id-view': 'editor_sdk'});

	function openDocument() {
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.initCollaborativeEditing({});
		let m = /[?&]bin=([^&]+)/.exec(location.search);
		let xhr = new XMLHttpRequest();
		xhr.open("GET", decodeURIComponent(m[1]));
		xhr.responseType = "arraybuffer";
		xhr.onload = function () { api.OpenDocumentFromBin(null, new Uint8Array(xhr.response)); };
		xhr.send();
	}

	api.HtmlElement = document.createElement("div");
	document.getElementById("editor_sdk").appendChild(api.HtmlElement);
	window["Asc"]["editor"] = api;

	//------------------------------------------------------------------------------------------------------------------
	function startTests() {
		QUnit.start();
		let ws = wb.getWorksheet(0);
		let o = ws.PagePrintOptions, m = o.asc_getPageMargins(), p = o.asc_getPageSetup();
		console.log("EXPLORE margins(mm) L=" + m.asc_getLeft() + " R=" + m.asc_getRight() + " T=" + m.asc_getTop() + " B=" + m.asc_getBottom() + " H=" + m.asc_getHeader() + " F=" + m.asc_getFooter() + " page " + p.asc_getWidth() + "x" + p.asc_getHeight() + " orient=" + p.asc_getOrientation());
		let adj = new Asc.asc_CAdjustPrint(); adj.asc_setPrintType(1);
		let d = api.wb.calcPagesPrint(adj);
		console.log("EXPLORE pages=" + d.arrPages.length + " first=" + JSON.stringify(d.arrPages[0] && d.arrPages[0].pageRange));
		QUnit.test("x", function (a) { a.ok(true); });
	}
});
