// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../../common/Drawings/WorkEvents.js'), 'utf8');
const begin = source.indexOf('// eo-macctrl:begin');
const end = source.indexOf('// eo-macctrl:end');
assert(begin >= 0 && end > begin);

function load(isMacOs)
{
	const AscCommon = {AscBrowser: {isMacOs}};
	const context = vm.createContext({AscCommon, global_mouseEvent: {ButtonOverride: -1}});
	vm.runInContext(source.slice(begin, end) + ';this.getMouseButton = getMouseButton;', context);
	return context;
}

const ev = (type, o) => Object.assign({type, button: 0, ctrlKey: false, metaKey: false, altKey: false, shiftKey: false}, o);

test('macOS: Control+click is a right click on press and on release', () =>
{
	const c = load(true);
	assert.equal(c.getMouseButton(ev('mousedown', {ctrlKey: true})), 2);
	const up = ev('mouseup', {ctrlKey: true});
	assert.equal(c.getMouseButton(up), 2);
	assert.equal(c.getMouseButton(up), 2, 'same release asked twice (cell + common handlers)');
});

test('macOS: Control released before the button still ends as a right click', () =>
{
	const c = load(true);
	assert.equal(c.getMouseButton(ev('pointerdown', {ctrlKey: true})), 2);
	assert.equal(c.getMouseButton(ev('pointerup')), 2);
});

test('macOS: a later unrelated release is a normal left release', () =>
{
	const c = load(true);
	c.getMouseButton(ev('mousedown', {ctrlKey: true}));
	c.getMouseButton(ev('mouseup', {ctrlKey: true}));
	assert.equal(c.getMouseButton(ev('mouseup')), 0);
	assert.equal(c.getMouseButton(ev('mousedown')), 0);
	assert.equal(c.getMouseButton(ev('mouseup')), 0);
});

test('macOS: plain, Cmd and Option clicks stay left clicks; real buttons unchanged', () =>
{
	const c = load(true);
	assert.equal(c.getMouseButton(ev('mousedown')), 0);
	assert.equal(c.getMouseButton(ev('mousedown', {altKey: true})), 0);
	assert.equal(c.getMouseButton(ev('mousedown', {metaKey: true})), 0);
	assert.equal(c.getMouseButton(ev('mousedown', {ctrlKey: true, metaKey: true})), 0);
	assert.equal(c.getMouseButton(ev('mousedown', {button: 2})), 2);
	assert.equal(c.getMouseButton(ev('mousedown', {button: 1})), 1);
});

test('Linux and Windows: Control+click is unchanged', () =>
{
	const c = load(false);
	assert.equal(c.getMouseButton(ev('mousedown', {ctrlKey: true})), 0);
	assert.equal(c.getMouseButton(ev('mouseup', {ctrlKey: true})), 0);
	assert.equal(c.getMouseButton(ev('mousedown', {button: 2})), 2);
});

test('events without a type (touch, synthetic) are not converted', () =>
{
	const c = load(true);
	assert.equal(c.getMouseButton({button: 0, ctrlKey: true}), 0);
});
