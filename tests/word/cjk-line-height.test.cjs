// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// #39: the line height of East Asian text in a substituted font follows Word's line height for the requested font.
// Run: node --test tests/word/cjk-line-height.test.cjs
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'),
	path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const styles = fs.readFileSync(path.join(root, 'word/Editor/Styles.js'), 'utf8');
const run = fs.readFileSync(path.join(root, 'word/Editor/Run.js'), 'utf8');

function slice(from, to) {
	const a = styles.indexOf(from);
	assert(a >= 0, from);
	const b = styles.indexOf(to, a);
	assert(b > a, to);
	return styles.slice(a, b + to.length);
}

function load(measurer) {
	const ctx = vm.createContext({AscWord: {}, g_oTextMeasurer: measurer});
	const factors = slice('var CJK_LINE_HEIGHT_FACTORS = {};', 'AscWord.GetCjkLineHeightFactor = GetCjkLineHeightFactor;');
	const update = slice('CTextMetrics.prototype.Update = function(fontName, fontSize, fontStyle, requestedName)', '\n};');
	vm.runInContext('function CTextMetrics(){this.Ascent=0;this.Descent=0;this.LineGap=0;this.Height=0;}\n' + factors + '\n' + update, ctx);
	return ctx;
}

const notoMeasurer = (height, ascent, descent) => ({
	SetFontInternal() {}, GetHeight: () => height, GetAscender: () => ascent, GetDescender: () => -descent
});

test('factors measured with Word for common fonts', () => {
	const ctx = load(notoMeasurer(1, 1, 0));
	const f = ctx.AscWord.GetCjkLineHeightFactor;
	assert.equal(f('SimSun'), 1.30);
	assert.equal(f('宋体'), 1.30);
	assert.equal(f('FangSong'), 1.30);
	assert.equal(f('新細明體'), 1.30);
	assert.equal(f('PMingLiU'), 1.30);
	assert.equal(f('DengXian'), 1.36);
	assert.equal(f('ＭＳ 明朝'), 1.30);
	assert.equal(f('MS Mincho'), 1.30);
	assert.equal(f('ＭＳ ゴシック'), 1.30);
	assert.equal(f('Yu Mincho'), 1.68);
	assert.equal(f('游ゴシック'), 1.68);
	assert.equal(f('Meiryo'), 1.95);
	assert.equal(f('ＭＳ Ｐ明朝'), 1.30);
	assert.equal(f('Microsoft YaHei'), 1.73);
	assert.equal(f('Arial'), 0);
	assert.equal(f(''), 0);
	assert.equal(f(undefined), 0);
});

test('Noto Sans CJK substituting SimSun gets Word\'s height, keeping the proportions', () => {
	// 12 pt Noto Sans CJK: 1.88 em = 7.96 mm; ascent 5.83 mm, descent 2.14 mm
	const ctx = load(notoMeasurer(7.967, 5.829, 2.138));
	const m = new ctx.CTextMetrics();
	m.Update('Noto Sans CJK SC', 12, 0, 'SimSun');
	const target = 1.30 * 12 * 25.4 / 72;
	assert.ok(Math.abs(m.Height - target) < 1e-9);
	assert.ok(Math.abs(m.Ascent / m.Descent - 5.829 / 2.138) < 1e-9);
});

test('other fonts and unlisted requests keep the font metrics', () => {
	const ctx = load(notoMeasurer(5, 4, 1));
	let m = new ctx.CTextMetrics();
	m.Update('Liberation Serif', 12, 0, 'SimSun');
	assert.equal(m.Height, 5);
	m = new ctx.CTextMetrics();
	m.Update('Noto Sans CJK SC', 12, 0, 'Some Unlisted Font');
	assert.equal(m.Height, 5);
	m = new ctx.CTextMetrics();
	m.Update('Noto Sans CJK SC', 12, 0);
	assert.equal(m.Height, 5);
});

test('the run passes the requested East Asian font to the metrics', () => {
	assert.ok(run.includes('metrics.Update(fontName, fontMap[fontId], fontStyle, eastAsiaName);'));
	assert.ok(run.includes('ParaRun.prototype.getRequestedEastAsianFontName = function(textPr)'));
});
