// SPDX-FileCopyrightText: 2026 ElixirOffice contributors
// SPDX-License-Identifier: AGPL-3.0-only
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../../common/clipboard_base.js'), 'utf8');

function declaredFunction(name)
{
	const marker = 'function ' + name + '(';
	const start = source.indexOf(marker);
	assert(start >= 0, 'missing ' + name);
	const brace = source.indexOf('{', start);
	let depth = 0;
	for (let i = brace; i < source.length; ++i)
	{
		if (source[i] === '{') ++depth;
		else if (source[i] === '}' && --depth === 0)
			return source.slice(start, i + 1);
	}
	assert.fail('unterminated ' + name);
}

function assignedMethod(name)
{
	const marker = name + ' : function(';
	const start = source.indexOf(marker);
	assert(start >= 0, 'missing ' + name);
	const functionStart = source.indexOf('function', start);
	const brace = source.indexOf('{', functionStart);
	let depth = 0;
	for (let i = brace; i < source.length; ++i)
	{
		if (source[i] === '{') ++depth;
		else if (source[i] === '}' && --depth === 0)
			return source.slice(functionStart, i + 1);
	}
	assert.fail('unterminated ' + name);
}

const context = vm.createContext({
	Node: {ELEMENT_NODE: 1, TEXT_NODE: 3, COMMENT_NODE: 8, DOCUMENT_FRAGMENT_NODE: 11}
});

function load(name)
{
	const result = vm.runInContext('(' + declaredFunction(name) + ')', context);
	context[name] = result;
	return result;
}

function element(name, children = [])
{
	return {nodeType: 1, nodeName: name.toUpperCase(), childNodes: children};
}

function text(value)
{
	return {nodeType: 3, nodeValue: value};
}

function fragment(children)
{
	return {nodeType: 11, childNodes: children};
}

test('image-only HTML is distinguished from mixed rich HTML', () =>
{
	const isImageOnly = load('isClipboardImageOnlyTree');
	assert.equal(isImageOnly(fragment([element('meta'), text('\n '), element('a', [element('img')])])), true);
	assert.equal(isImageOnly(fragment([element('picture', [element('source'), element('img')])])), true);
	assert.equal(isImageOnly(fragment([element('p', [element('img'), element('br')])])), true);
	assert.equal(isImageOnly(fragment([element('p', [text('caption'), element('img')])])), false);
	assert.equal(isImageOnly(fragment([element('table', [element('img')])])), false);
	assert.equal(isImageOnly(fragment([text('   ')])), false);
});

test('representative Windows, macOS and Linux flavour sets expose bitmap files', () =>
{
	const getImages = load('getClipboardImageFiles');
	const bitmap = type => ({kind: 'file', type, getAsFile: () => ({type})});
	const html = {kind: 'string', type: 'text/html'};
	const plain = {kind: 'string', type: 'text/plain'};
	const sets = [
		[html, plain, bitmap('image/png')],
		[html, plain, bitmap('image/tiff')],
		[html, {kind: 'string', type: 'text/uri-list'}, bitmap('image/png')]
	];
	const actual = sets.map(items => getImages(items).map(file => file.type));
	assert.deepEqual(JSON.parse(JSON.stringify(actual)), [
		['image/png'], ['image/tiff'], ['image/png']
	]);
	assert.equal(getImages([html, {kind: 'file', type: 'application/pdf', getAsFile: () => ({})}]).length, 0);
});

test('paste routing prefers bitmap only for image-only HTML', () =>
{
	const shouldPrefer = load('shouldPreferClipboardImages');
	context.isClipboardImageOnlyTree = load('isClipboardImageOnlyTree');
	const imageOnly = fragment([element('div', [element('img')])]);
	const mixed = fragment([element('div', [text('hello'), element('img')])]);
	assert.equal(shouldPrefer(imageOnly, [{type: 'image/png'}]), true);
	assert.equal(shouldPrefer(mixed, [{type: 'image/png'}]), false);
	assert.equal(shouldPrefer(imageOnly, []), false);

	assert.match(source, /shouldPreferClipboardImages\([^)]*\)[\s\S]*readClipboardImagesAsHtml/);
	assert.match(source, /CommonIframe_PasteStart\(_html_format, _text_format\)/);
	assert.match(source, /reader\.readAsDataURL\(blob\)/);
});

test('bitmap files are converted to local image HTML without a remote URL', () =>
{
	context.FileReader = class
	{
		readAsDataURL(image)
		{
			if (image.error)
				this.onerror({target: this});
			else
				this.onload({target: {result: image.data}});
		}
	};
	const readImages = load('readClipboardImagesAsHtml');
	let actual;
	readImages([
		{data: 'data:image/png;base64,AAAA'},
		{error: true},
		{data: 'data:image/jpeg;base64,BBBB'}
	], html => { actual = html; });
	assert.equal(actual, '<html><body><img src="data:image/png;base64,AAAA"/><img src="data:image/jpeg;base64,BBBB"/></body></html>');
	assert.doesNotMatch(actual, /https?:/);
});

test('the real paste handler selects bitmap, rich HTML and safe fallbacks', () =>
{
	context.isClipboardImageOnlyTree = load('isClipboardImageOnlyTree');
	context.getClipboardImageFiles = load('getClipboardImageFiles');
	context.shouldPreferClipboardImages = load('shouldPreferClipboardImages');
	context.readClipboardImagesAsHtml = load('readClipboardImagesAsHtml');
	context.FileReader = class
	{
		readAsDataURL(image)
		{
			this.onload({target: {result: image.data}});
		}
	};
	context.AscCommon = {
		c_oAscClipboardDataFormat: {HtmlElement: 8, Internal: 4, Text: 1},
		g_specialPasteHelper: {specialPasteData: {}}
	};
	const handler = vm.runInContext('(' + assignedMethod('_private_onpaste') + ')', context);
	const bitmapItem = {
		kind: 'file',
		type: 'image/png',
		getAsFile: () => ({data: 'data:image/png;base64,PIXELS'})
	};

	function run({html = '', plain = '', items = [], tree})
	{
		const calls = [];
		context.document = {
			createElement: () => ({
				content: tree,
				set innerHTML(value) { this.value = value; }
			})
		};
		const receiver = {
			Api: {
				asc_IsFocus: () => true,
				isLongAction: () => false,
				incrementCounterLongAction() {}
			},
			IsNeedDivOnPaste: false,
			ClosureParams: {
				_e: null,
				getData: type => type === 'text/html' ? html : type === 'text/plain' ? plain : ''
			},
			CommonIframe_PasteStart: (...args) => calls.push(args),
			Paste_End() {},
			_console_log() {},
			PasteImagesBody: '',
			PasteImagesCount: 0,
			PasteImagesCounter: 0
		};
		context.g_clipboardBase = receiver;
		handler.call(receiver, {clipboardData: {items, getData() {}}, preventDefault() {}});
		return calls;
	}

	const imageOnly = fragment([element('img')]);
	const mixed = fragment([text('caption'), element('img')]);
	assert.deepEqual(Array.from(run({html: '<img src="https://remote/image">', items: [bitmapItem], tree: imageOnly})[0]),
		['<html><body><img src="data:image/png;base64,PIXELS"/></body></html>']);
	assert.deepEqual(Array.from(run({html: '<p>caption<img src="https://remote/image"></p>', plain: 'caption', items: [bitmapItem], tree: mixed})[0]),
		['<p>caption<img src="https://remote/image"></p>', 'caption']);
	assert.deepEqual(Array.from(run({html: '<img src="https://remote/image">', tree: imageOnly})[0]),
		['<img src="https://remote/image">', '']);
	assert.deepEqual(Array.from(run({items: [bitmapItem]})[0]),
		['<html><body><img src="data:image/png;base64,PIXELS"/></body></html>']);
});
