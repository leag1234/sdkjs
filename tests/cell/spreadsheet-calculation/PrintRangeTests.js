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
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	}

	api.HtmlElement = document.createElement("div");
	document.getElementById("editor_sdk").appendChild(api.HtmlElement);
	window["Asc"]["editor"] = api;

	//------------------------------------------------------------------------------------------------------------------
	const PrintType = Asc.c_oAscPrintType;
	let sheetCounter = 0;

	// creates a new empty sheet (model + view) and returns its index
	function addSheet() {
		let index = wb.getWorksheetCount();
		wb.createWorksheet(index, "P" + (++sheetCounter));
		return index;
	}

	function fillContent(ws, cols, rows) {
		for (let r = 0; r < rows; ++r) {
			for (let c = 0; c < cols; ++c) {
				ws.getRange3(r, c, r, c).setValue("v" + r + "_" + c);
			}
		}
	}

	function adjustFor(printType, sheets, ignorePrintArea) {
		let adjust = new Asc.asc_CAdjustPrint();
		adjust.asc_setPrintType(printType);
		if (sheets) {
			adjust.asc_setActiveSheetsArray(sheets);
		}
		if (undefined !== ignorePrintArea) {
			adjust.asc_setIgnorePrintArea(ignorePrintArea);
		}
		return adjust;
	}

	function pagesOfSheet(index, ignorePrintArea) {
		let data = api.wb.calcPagesPrint(adjustFor(PrintType.ActiveSheets, [index], ignorePrintArea));
		return data.arrPages.filter(function (page) {
			return page.indexWorksheet === index;
		});
	}

	function lastCell(pages) {
		let res = {col: -1, row: -1};
		for (let i = 0; i < pages.length; ++i) {
			res.col = Math.max(res.col, pages[i].pageRange.c2);
			res.row = Math.max(res.row, pages[i].pageRange.r2);
		}
		return res;
	}

	function setPrintArea(index, ref) {
		let ws = wb.getWorksheet(index);
		wb.addDefName("Print_Area", "'" + ws.getName() + "'!" + ref, ws.getId(), false, false);
	}

	function setScale(index, scale, fitToWidth, fitToHeight) {
		let ws = wb.getWorksheet(index);
		let pageSetup = ws.PagePrintOptions.asc_getPageSetup();
		ws.setFitToPage(!!(fitToWidth || fitToHeight));
		pageSetup.asc_setScale(scale, true);
		pageSetup.asc_setFitToWidth(fitToWidth || 0, true);
		pageSetup.asc_setFitToHeight(fitToHeight || 0, true);
	}

	function startTests() {
		QUnit.start();
		QUnit.module("Print range");

		QUnit.test("Width-only columns do not extend the print range", function (assert) {
			let plain = addSheet();
			fillContent(wb.getWorksheet(plain), 14, 15);
			let widthOnly = addSheet();
			let ws = wb.getWorksheet(widthOnly);
			fillContent(ws, 14, 15);
			ws.setColWidth(12, 14, 2999);
			ws.setRowHeight(25, 15, 400, true);
			assert.ok(ws.getColsCount() > 2900, "the sheet declares thousands of columns");

			let plainPages = pagesOfSheet(plain);
			let pages = pagesOfSheet(widthOnly);
			assert.deepEqual(lastCell(pages), {col: 13, row: 14}, "print range ends at the last cell with a value");
			assert.strictEqual(pages.length, plainPages.length, "same page count as the sheet without width-only columns");

			let info = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [widthOnly]))[0];
			assert.strictEqual(info.lastCol, 13, "range info: last column");
			assert.strictEqual(info.lastRow, 14, "range info: last row");
			assert.strictEqual(info.sources.cells.col, 13, "range info: last column comes from cells");
			assert.ok(info.colsCount > 2900, "range info: reports the declared columns count");
		});

		QUnit.test("Column default style without a visible format does not extend the print range", function (assert) {
			let index = addSheet();
			let ws = wb.getWorksheet(index);
			fillContent(ws, 14, 15);
			// whole-column number format and font: no fill, no border
			let columns = ws.getRange3(0, 14, AscCommon.gc_nMaxRow0, 2999);
			columns.setNumFormat("0.00");
			columns.setBold(true);
			assert.deepEqual(lastCell(pagesOfSheet(index)), {col: 13, row: 14}, "whole-column style is ignored");

			// empty cells that only carry a number format / bold font
			let styled = addSheet();
			let ws2 = wb.getWorksheet(styled);
			fillContent(ws2, 14, 15);
			let block = ws2.getRange3(0, 14, 14, 200);
			block.setNumFormat("0.00");
			block.setBold(true);
			assert.deepEqual(lastCell(pagesOfSheet(styled)), {col: 13, row: 14}, "style-only empty cells are ignored");
		});

		QUnit.test("Formatted empty cells still extend the print range", function (assert) {
			let bordered = addSheet();
			let ws = wb.getWorksheet(bordered);
			fillContent(ws, 14, 15);
			let border = new AscCommonExcel.Border();
			["l", "r", "t", "b"].forEach(function (side) {
				border[side] = new AscCommonExcel.BorderProp();
				border[side].setStyle(Asc.c_oAscBorderStyles.Thin);
				border[side].c = AscCommonExcel.createRgbColor(0, 0, 0);
			});
			ws.getRange3(19, 19, 19, 19).setBorder(border);
			assert.deepEqual(lastCell(pagesOfSheet(bordered)), {col: 19, row: 19}, "an empty bordered cell extends the range");

			let filled = addSheet();
			let ws2 = wb.getWorksheet(filled);
			fillContent(ws2, 14, 15);
			ws2.getRange3(24, 17, 24, 17).setFillColor(AscCommonExcel.createRgbColor(255, 255, 0));
			assert.deepEqual(lastCell(pagesOfSheet(filled)), {col: 17, row: 24}, "an empty filled cell extends the range");
			let info = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [filled]))[0];
			assert.strictEqual(info.lastCol, 17, "range info: last column");
			assert.strictEqual(info.lastRow, 24, "range info: last row");
		});

		QUnit.test("Merged cells extend the range only when they touch the printed area", function (assert) {
			let index = addSheet();
			let ws = wb.getWorksheet(index);
			fillContent(ws, 2, 5);
			ws.getRange3(0, 6, 1, 7).merge(Asc.c_oAscMergeOptions.Merge);
			assert.deepEqual(lastCell(pagesOfSheet(index)), {col: 7, row: 4}, "merged range counts for the whole sheet");

			setPrintArea(index, "$A:$B");
			let pages = pagesOfSheet(index, false);
			assert.ok(pages.length > 0, "the print area is printed");
			assert.strictEqual(lastCell(pages).col, 1, "a merged range outside of the whole-column print area is not printed");
		});

		QUnit.test("An explicit print area is honoured unless it is ignored", function (assert) {
			let index = addSheet();
			let ws = wb.getWorksheet(index);
			fillContent(ws, 14, 15);
			ws.setColWidth(12, 14, 2999);
			setPrintArea(index, "$A$1:$C$5");

			let pages = pagesOfSheet(index, false);
			assert.deepEqual(lastCell(pages), {col: 2, row: 4}, "print area A1:C5 is printed");
			assert.strictEqual(pages.length, 1, "one page for the print area");

			pages = pagesOfSheet(index, true);
			assert.deepEqual(lastCell(pages), {col: 13, row: 14}, "ignoring the print area prints the used range, not the width-only columns");

			// the print area may be larger than the content: it is printed as defined
			let wide = addSheet();
			let ws2 = wb.getWorksheet(wide);
			fillContent(ws2, 3, 3);
			setPrintArea(wide, "$A$1:$H$30");
			assert.deepEqual(lastCell(pagesOfSheet(wide, false)), {col: 7, row: 29}, "print area larger than the content");
		});

		QUnit.test("EntireWorkbook prints every visible sheet with content and skips empty ones", function (assert) {
			let first = addSheet();
			let empty = addSheet();
			let second = addSheet();
			let hidden = addSheet();
			let lastEmpty = addSheet();
			fillContent(wb.getWorksheet(first), 3, 3);
			fillContent(wb.getWorksheet(second), 2, 2);
			fillContent(wb.getWorksheet(hidden), 2, 2);
			wb.getWorksheet(hidden).setHidden(true);

			let data = api.wb.calcPagesPrint(adjustFor(PrintType.EntireWorkbook));
			let sheets = {};
			data.arrPages.forEach(function (page) {
				sheets[page.indexWorksheet] = (sheets[page.indexWorksheet] || 0) + 1;
			});
			assert.strictEqual(sheets[first], 1, "first sheet is printed");
			assert.strictEqual(sheets[second], 1, "second sheet is printed");
			assert.strictEqual(sheets[empty], undefined, "empty sheet is skipped");
			assert.strictEqual(sheets[lastEmpty], undefined, "trailing empty sheet is skipped");
			assert.strictEqual(sheets[hidden], undefined, "hidden sheet is skipped");

			data = api.wb.calcPagesPrint(adjustFor(PrintType.ActiveSheets, [first]));
			assert.strictEqual(data.arrPages.length, 1, "ActiveSheets prints only the requested sheet");
			assert.strictEqual(data.arrPages[0].indexWorksheet, first, "ActiveSheets: the requested sheet");
		});

		QUnit.test("Page margins that leave no printable area fall back to the defaults", function (assert) {
			function setMargins(index, left, right, top, bottom) {
				let margins = wb.getWorksheet(index).PagePrintOptions.pageMargins;
				margins.left = left;
				margins.right = right;
				margins.top = top;
				margins.bottom = bottom;
			}

			let normal = addSheet();
			let broken = addSheet();
			let large = addSheet();
			[normal, broken, large].forEach(function (index) {
				fillContent(wb.getWorksheet(index), 3, 60);
				setScale(index, 100);
			});
			let defaultPages = pagesOfSheet(normal);
			assert.ok(defaultPages.length > 1 && defaultPages.length < 6, "default margins: a few pages for 60 rows");

			// bottom margin of 149 cm and a 90 cm footer on an A4 page (seen in a real ods file)
			setMargins(broken, 17.8, 17.8, 19.1, 1490);
			let pages = pagesOfSheet(broken);
			assert.strictEqual(pages.length, defaultPages.length, "broken margins: same pagination as the defaults");
			assert.ok(pages[0].pageRange.r2 > 5, "rows are not printed one per page");
			assert.strictEqual(pages[0].pageClipRectHeight, defaultPages[0].pageClipRectHeight, "printable height of the defaults");

			// large but usable margins are kept
			setMargins(large, 17.8, 17.8, 100, 100);
			assert.ok(pagesOfSheet(large).length > defaultPages.length, "usable margins are respected");

			// horizontal margins wider than the page
			let wide = addSheet();
			fillContent(wb.getWorksheet(wide), 3, 10);
			setMargins(wide, 300, 300, 19.1, 19.1);
			assert.strictEqual(pagesOfSheet(wide).length, 1, "margins wider than the page: defaults are used");

			// reporting
			let info = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [broken]))[0];
			assert.strictEqual(info.marginsFallback, true, "range info: fallback reported");
			assert.strictEqual(info.margins.bottom, 1490, "range info: stored margin");
			assert.strictEqual(info.usedMargins.bottom, 19.1, "range info: used margin");
			assert.strictEqual(api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [normal]))[0].marginsFallback, false, "range info: no fallback");

			let warnings = api.asc_GetPrintSetupWarnings(adjustFor(PrintType.ActiveSheets, [broken]));
			assert.deepEqual(warnings.map(function (w) { return w.kind; }), ["margins"], "setup warnings: margins");
			assert.strictEqual(warnings[0].sheetIndex, broken, "setup warnings: sheet");
			assert.deepEqual(api.asc_GetPrintSetupWarnings(adjustFor(PrintType.ActiveSheets, [normal])), [], "setup warnings: none for a normal sheet");
			assert.deepEqual(api.asc_GetPrintSetupWarnings(adjustFor(PrintType.ActiveSheets, [large])), [], "setup warnings: usable large margins");

			// tiny scale and broken margins are both reported
			setScale(broken, 10);
			warnings = api.asc_GetPrintSetupWarnings(adjustFor(PrintType.ActiveSheets, [broken]));
			assert.deepEqual(warnings.map(function (w) { return w.kind; }), ["scale", "margins"], "setup warnings: scale and margins");
			assert.deepEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [broken])).length, 1, "scale warnings keep their shape");
		});

		QUnit.test("asc_GetPrintRangeInfo reports the printed size and the page setup", function (assert) {
			let index = addSheet();
			let ws = wb.getWorksheet(index);
			fillContent(ws, 14, 15);
			ws.setColWidth(12, 14, 2999);
			setScale(index, 80, 0, 0);
			let wsView = api.wb.getWorksheet(index);
			let info = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [index]))[0];

			let expectedWidth = 0;
			for (let c = 0; c < 14; ++c) {
				expectedWidth += (wsView._getWidthForPrint(c) || wsView._getColumnWidth(c)) * AscCommonExcel.vector_koef / wsView.getZoom();
			}
			assert.ok(info.widthMm > 14 * 10 && info.widthMm < 14 * 60, "width is the sum of 14 column widths (" + info.widthMm.toFixed(1) + " mm)");
			assert.ok(Math.abs(info.widthMm - expectedWidth) < 0.5, "width comes from the print column widths");
			assert.ok(info.heightMm > 15 * 3 && info.heightMm < 15 * 15, "height is the sum of 15 row heights (" + info.heightMm.toFixed(1) + " mm)");
			assert.strictEqual(info.hasPrintArea, false, "no print area");
			assert.strictEqual(info.pageWidthMm, 210, "page width");
			assert.strictEqual(info.pageHeightMm, 297, "page height");
			assert.ok(Math.abs(info.printableWidthMm - (210 - 17.8 * 2)) < 0.01, "printable width");
			assert.ok(Math.abs(info.printableHeightMm - (297 - 19.1 * 2)) < 0.01, "printable height");
			assert.strictEqual(info.orientation, "portrait", "orientation");
			assert.strictEqual(info.scale, 80, "stored scale");
			assert.strictEqual(info.fitToWidth, 0, "fitToWidth");
			assert.strictEqual(info.fitToHeight, 0, "fitToHeight");
			assert.strictEqual(info.horizontalCentered, false, "horizontalCentered");

			// print area A1:C5, honoured unless ignored
			setPrintArea(index, "$A$1:$C$5");
			let withArea = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [index], false))[0];
			assert.strictEqual(withArea.hasPrintArea, true, "print area applies");
			assert.ok(Math.abs(withArea.widthMm - info.widthMm * 3 / 14) < info.widthMm / 14 * 0.6, "width of three columns");
			assert.ok(Math.abs(withArea.heightMm - info.heightMm * 5 / 15) < 3, "height of five rows");
			let ignored = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [index], true))[0];
			assert.strictEqual(ignored.hasPrintArea, false, "ignored print area");
			assert.ok(Math.abs(ignored.widthMm - info.widthMm) < 0.01, "ignored: used range");

			// landscape, fit to width and centering from the print settings
			let options = ws.PagePrintOptions.clone(ws);
			options.pageSetup.orientation = Asc.c_oAscPageOrientation.PageLandscape;
			options.horizontalCentered = true;
			let adjust = adjustFor(PrintType.ActiveSheets, [index], true);
			adjust.pageOptionsMap = [];
			adjust.pageOptionsMap[index] = options;
			let custom = api.asc_GetPrintRangeInfo(adjust)[0];
			assert.strictEqual(custom.orientation, "landscape", "landscape from the print settings");
			assert.strictEqual(custom.pageWidthMm, 297, "landscape page width");
			assert.ok(Math.abs(custom.printableWidthMm - (297 - 17.8 * 2)) < 0.01, "landscape printable width");
			assert.strictEqual(custom.horizontalCentered, true, "horizontalCentered from the print settings");

			// fields exist for an empty sheet too
			let empty = api.asc_GetPrintRangeInfo(adjustFor(PrintType.ActiveSheets, [addSheet()]))[0];
			assert.strictEqual(empty.widthMm, 0, "empty sheet: no width");
			assert.strictEqual(empty.heightMm, 0, "empty sheet: no height");
		});

		QUnit.test("asc_GetPrintScaleWarnings", function (assert) {
			let tiny = addSheet();
			let normal = addSheet();
			let fit = addSheet();
			let emptyTiny = addSheet();
			let hiddenTiny = addSheet();
			[tiny, normal, fit, hiddenTiny].forEach(function (index) {
				fillContent(wb.getWorksheet(index), 3, 3);
			});
			setScale(tiny, 10);
			setScale(normal, 100);
			setScale(fit, 10, 1, 0);
			setScale(emptyTiny, 10);
			setScale(hiddenTiny, 10);
			wb.getWorksheet(hiddenTiny).setHidden(true);

			let warnings = api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [tiny]));
			assert.deepEqual(warnings, [{sheetIndex: tiny, name: wb.getWorksheet(tiny).getName(), scale: 10}], "tiny scale is reported");

			assert.deepEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [normal])), [], "100% is not reported");
			assert.deepEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [fit])), [], "fit to width: the stored scale is ignored");
			assert.deepEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [emptyTiny])), [], "empty sheet is not reported");
			assert.deepEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [hiddenTiny])), [], "hidden sheet is not reported");

			setScale(normal, 24);
			assert.strictEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [normal])).length, 1, "24% is reported");
			setScale(normal, 25);
			assert.strictEqual(api.asc_GetPrintScaleWarnings(adjustFor(PrintType.ActiveSheets, [normal])).length, 0, "25% is not reported");

			warnings = api.asc_GetPrintScaleWarnings(adjustFor(PrintType.EntireWorkbook));
			let reported = warnings.map(function (w) { return w.sheetIndex; });
			assert.ok(reported.indexOf(tiny) !== -1, "EntireWorkbook: tiny sheet is reported");
			assert.ok(reported.indexOf(fit) === -1 && reported.indexOf(emptyTiny) === -1 && reported.indexOf(hiddenTiny) === -1,
				"EntireWorkbook: fit, empty and hidden sheets are not reported");

			// per-sheet options passed with the print settings override the stored ones
			let adjust = adjustFor(PrintType.ActiveSheets, [normal]);
			let options = wb.getWorksheet(normal).PagePrintOptions.clone(wb.getWorksheet(normal));
			options.asc_getPageSetup().asc_setScale(10, true);
			adjust.pageOptionsMap = [];
			adjust.pageOptionsMap[normal] = options;
			assert.strictEqual(api.asc_GetPrintScaleWarnings(adjust).length, 1, "pageOptions of the print settings are used");

			assert.deepEqual(api.asc_GetPrintScaleWarnings(), [], "default settings: the active sheet is empty");
		});
	}
});
