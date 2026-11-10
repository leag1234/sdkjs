"use strict";

// Character borders (w:bdr, GitHub #66): adjacent runs with an equal border are drawn as one box
$(function () {
	
	let logicDocument = AscTest.CreateLogicDocument();
	
	const charWidth = AscTest.CharWidth * AscTest.FontSize;
	
	function initDocument(widthInChars)
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, AscTest.CreateParagraph());
		
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize((widthInChars + 10) * charWidth, 1000);
		sectPr.SetPageMargins(5 * charWidth, 50, 5 * charWidth, 50);
	}
	
	function createBorder(size, space)
	{
		let border = new CDocumentBorder();
		border.Value = border_Single;
		border.Size  = size;
		border.Space = space;
		border.Color = new CDocumentColor(255, 0, 0, false);
		return border;
	}
	
	function addRun(p, text, border)
	{
		let run = AscTest.AddTextToParagraph(p, text);
		if (border)
		{
			run.Pr.Border = border;
			run.Recalc_CompiledPr(true);
		}
		return run;
	}
	
	// Draws the paragraph into a recording graphics and returns the lines of the boxes
	function drawBorders(p)
	{
		let calls = [];
		let line  = -1;
		let g = new Proxy({}, {
			get : function(target, name)
			{
				if ("Start_Command" === name)
					return function(cmd, l, curLine) { line = curLine; };
				if ("drawHorLine" === name)
					return function(rule, y, x0, x1, w) { calls.push({type : "hor", line : line, y : y, x0 : x0, x1 : x1, w : w}); };
				if ("drawVerLine" === name)
					return function(rule, x, y0, y1, w) { calls.push({type : "ver", line : line, x : x, y0 : y0, y1 : y1, w : w}); };
				if ("string" !== typeof name || name.startsWith("m_") || "isPrintMode" === name || "IsThumbnail" === name || "then" === name)
					return undefined;
				return function() { return 0; };
			}
		});
		
		let drawState = new AscWord.ParagraphDrawState();
		drawState.init(p, g);
		p.drawRunContentLines(0, g, drawState);
		return calls;
	}
	
	QUnit.module("Test character borders", {
		beforeEach : function()
		{
			initDocument(60);
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		}
	});
	
	QUnit.test("Two adjacent runs with an equal border make one box", function(assert)
	{
		let p = logicDocument.GetElement(0);
		let x0 = 5 * charWidth;
		
		addRun(p, "aaa", null);
		addRun(p, "bbb", createBorder(0.5, 1));
		addRun(p, "ccc", createBorder(0.5, 1));
		addRun(p, "ddd", null);
		AscTest.Recalculate();
		
		let calls = drawBorders(p);
		let hor   = calls.filter(c => "hor" === c.type);
		let ver   = calls.filter(c => "ver" === c.type);
		
		assert.strictEqual(hor.length, 2, "One top and one bottom line for the merged box");
		assert.close(hor[0].x0, x0 + 3 * charWidth - 1, 0.001, "The box starts at the first bordered run (minus the space)");
		assert.close(hor[0].x1, x0 + 9 * charWidth + 1, 0.001, "The box ends at the end of the second bordered run (plus the space)");
		assert.strictEqual(ver.length, 2, "Only the outer edges are drawn, nothing between the runs");
		assert.close(Math.min(ver[0].x, ver[1].x), x0 + 3 * charWidth - 1, 0.001, "Left edge at the first run");
		assert.close(Math.max(ver[0].x, ver[1].x), x0 + 9 * charWidth + 1, 0.001, "Right edge at the last run");
	});
	
	QUnit.test("Different borders make separate boxes", function(assert)
	{
		let p = logicDocument.GetElement(0);
		
		addRun(p, "aaa", createBorder(0.5, 0));
		addRun(p, "bbb", createBorder(1.0, 0));
		AscTest.Recalculate();
		
		let calls = drawBorders(p);
		assert.strictEqual(calls.filter(c => "hor" === c.type).length, 4, "Two boxes: two horizontal lines each");
		assert.strictEqual(calls.filter(c => "ver" === c.type).length, 4, "Two boxes: two vertical lines each");
	});
	
	QUnit.test("A wrapped stretch has the left edge on the first line and the right edge on the last", function(assert)
	{
		initDocument(20);
		let p = logicDocument.GetElement(0);
		
		addRun(p, "one ", null);
		addRun(p, "two three four five ", createBorder(0.5, 0));
		addRun(p, "six seven eight nine ten eleven", createBorder(0.5, 0));
		addRun(p, " end", null);
		AscTest.Recalculate();
		
		assert.ok(p.getLineCount() >= 3, "The text wraps to at least 3 lines");
		
		let calls = drawBorders(p);
		let ver   = calls.filter(c => "ver" === c.type);
		let lines = Array.from(new Set(calls.map(c => c.line))).sort();
		let first = lines[0], last = lines[lines.length - 1];
		
		assert.strictEqual(lines.length, p.getLineCount(), "Every line has a box part");
		assert.strictEqual(ver.filter(c => c.line === first).length, 1, "First line: left edge only");
		assert.ok(ver.find(c => c.line === first).x < 5 * charWidth + 10 * charWidth, "It is the left edge");
		assert.strictEqual(ver.filter(c => c.line === last).length, 1, "Last line: right edge only");
		for (let l = first + 1; l < last; ++l)
			assert.strictEqual(ver.filter(c => c.line === l).length, 0, "Middle line " + l + ": no vertical edges");
	});
});
