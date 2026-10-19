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

// Execute the actual table-property implementation. Unrelated drawing constructors
// are inert here; file-wire and editor behavior are exercised by the engine matrix.
const source = fs.readFileSync(path.join(__dirname, '../../word/Editor/Styles.js'), 'utf8');
const start = source.indexOf('function CTablePr()');
const end = source.indexOf('function CTableRowHeight(', start);
assert(start > 0 && end > start);
const context = vm.createContext({
	AscCommon: {document_compatibility_mode_Word12: 12}, AscWord: {},
	align_Left: 1, tblwidth_Mm: 1, tblwidth_Auto: 0, tbllayout_AutoFit: 0,
	g_dKoef_pt_to_mm: 25.4 / 72,
	CDocumentShd: function () {},
	CDocumentBorder: function () {},
	CTableMeasurement: function (type, width) { this.Type = type; this.W = width; },
	IsEqualStyleObjects: (a, b) => a === b
});
vm.runInContext(source.slice(start, end), context);
const TablePr = context.CTablePr;

class Binary {
	constructor() { this.bytes = Buffer.alloc(4096); this.pos = 0; }
	GetCurPosition() { return this.pos; }
	Skip(n) { this.pos += n; }
	Seek(n) { this.pos = n; }
	WriteLong(n) { this.bytes.writeInt32LE(n, this.pos); this.pos += 4; }
	GetLong() { const n = this.bytes.readInt32LE(this.pos); this.pos += 4; return n; }
	WriteBool(n) { this.bytes[this.pos++] = n ? 1 : 0; }
	GetBool() { return !!this.bytes[this.pos++]; }
	WriteString2(s) { this.WriteLong(s.length); this.bytes.write(s, this.pos, 'utf16le'); this.pos += s.length * 2; }
	GetString2() { const n = this.GetLong() * 2; const s = this.bytes.toString('utf16le', this.pos, this.pos + n); this.pos += n; return s; }
}

test('copy and from-object retain true, explicit false and absence', () => {
	for (const direction of [true, false, undefined]) {
		const pr = new TablePr(); pr.BidiVisual = direction;
		assert.equal(pr.Copy().BidiVisual, direction);
		const copy = new TablePr(); copy.Set_FromObject(pr);
		assert.equal(copy.BidiVisual, direction);
	}
});
test('explicit false overrides inherited RTL; absence leaves inherited value', () => {
	const inherited = new TablePr(); inherited.BidiVisual = true;
	const absent = new TablePr(); inherited.Merge(absent);
	assert.equal(inherited.BidiVisual, true);
	const override = new TablePr(); override.BidiVisual = false; inherited.Merge(override);
	assert.equal(inherited.BidiVisual, false);
});
test('equality distinguishes direct direction and defaults preserve absence', () => {
	const a = new TablePr(), b = new TablePr();
	a.BidiVisual = true; b.BidiVisual = false;
	assert.equal(a.Is_Equal(b), false);
	const defaults = new TablePr(); defaults.InitDefault();
	assert.equal(defaults.BidiVisual, undefined);
});
test('internal binary preserves three states and the next record boundary', () => {
	for (const direction of [true, false, undefined]) {
		const pr = new TablePr(); pr.BidiVisual = direction; pr.TableCaption = 'table';
		const binary = new Binary(); pr.Write_ToBinary(binary);
		const boundary = binary.GetCurPosition(); binary.WriteLong(0x12345678);
		binary.Seek(0); const copy = new TablePr(); copy.Read_FromBinary(binary);
		assert.equal(copy.BidiVisual, direction);
		assert.equal(copy.TableCaption, 'table');
		assert.equal(binary.GetCurPosition(), boundary);
		assert.equal(binary.GetLong(), 0x12345678);
	}
});
test('packed boolean adds no payload an older unframed reader could misread', () => {
	const lengths = [];
	for (const direction of [true, false, undefined]) {
		const pr = new TablePr(); pr.BidiVisual = direction; pr.TableCaption = 'table';
		const binary = new Binary(); pr.Write_ToBinary(binary); lengths.push(binary.GetCurPosition());
	}
	assert.equal(new Set(lengths).size, 1);
});

test('implicit default resets to absence and compares equally without erasing direct overrides', () => {
	for (const prior of [true, false, undefined]) {
		const defaults = new TablePr(); defaults.BidiVisual = prior; defaults.InitDefault();
		assert.equal(defaults.BidiVisual, undefined, 'Defaults must not manufacture a direct false');
		const absent = Object.assign(new TablePr(), defaults); absent.BidiVisual = undefined;
		assert.equal(defaults.Is_Equal(absent), true, 'An implicit default equals absence');
	}
	const direct = new TablePr(), inherited = new TablePr();
	direct.BidiVisual = false;
	assert.equal(direct.Is_Equal(inherited), false, 'A direct false can override inherited RTL');
});
