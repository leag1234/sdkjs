// SPDX-License-Identifier: AGPL-3.0-only
'use strict';
// AscCommon.getLocalFileHyperlink: the links to files of this computer that the hyperlink dialogs of the editors accept (issue #77).
const {test} = require('node:test'), assert = require('node:assert/strict'), fs = require('node:fs'), path = require('node:path'), vm = require('node:vm');
const root = path.resolve(__dirname, '../..');
const src = fs.readFileSync(path.join(root, 'common/editorscommon.js'), 'utf8');
function slice(from, to) {
	const start = src.indexOf(from);
	assert(start >= 0, from);
	const end = src.indexOf(to, start);
	assert(end > start, to);
	return src.slice(start, end);
}
const ctx = vm.createContext({});
vm.runInContext(fs.readFileSync(path.join(root, 'vendor/string.js'), 'utf8').match(/String\.prototype\.strongMatch = function[\s\S]*?\n\t\};/)[0], ctx);
vm.runInContext('var fileRe = ' + src.match(/fileRe\s*=\s*(\/.*?\/i),/)[1] + ';' + slice('function isValidFileUrl', 'function isSafeHyperlink') +
	slice('function getLocalFileHyperlink', 'function getUrlType') + ';this.get = getLocalFileHyperlink; this.is = isLocalFileHyperlink;', ctx);
const plain = v => JSON.parse(JSON.stringify(v));

test('file: URLs, absolute paths of the three systems and relative paths are links to local files', () => {
	const cases = [
		['file:///Users/x/doc.pdf', 'url'], ['file:///C:/x/doc.pdf', 'url'], ['file://localhost/etc/a.pdf', 'url'], ['FILE:///Users/x/a b.pdf', 'url'],
		['/Users/x/doc.pdf', 'posix'], ['/home/user/123.txt', 'posix'], ['/tmp/a b.pdf', 'posix'],
		['C:\\x\\doc.pdf', 'windows'], ['c:/x/doc.pdf', 'windows'],
		['\\\\server\\share\\a.pdf', 'unc'],
		['..\\docs\\a.pdf', 'relative'], ['../docs/a.pdf', 'relative'], ['./a.pdf', 'relative'], ['.\\a.pdf', 'relative'],
		['  /Users/x/doc.pdf  ', 'posix']
	];
	for (const [url, kind] of cases) {
		assert.equal(ctx.get(url).kind, kind, url);
		assert.equal(ctx.is(url), true, url);
	}
});
test('everything else is not', () => {
	for (const url of ['https://example.org/a', 'http://a.b', 'mailto:a@b.cd', 'www.example.org', '//host/a', '/', '/ ', 'C:', 'C:x', 'a.pdf', 'docs/a.pdf',
		'smb://192.0.2.1/e/a.docx', 'javascript:alert(1)', 'file://"C:\\Users\\a.pdf"', 'file://', '', '   ', null, undefined, 12, '/tmp/a\n.pdf', '/tmp/a\u0000.pdf',
		'/' + 'a'.repeat(2100)])
		assert.equal(ctx.is(url), false, JSON.stringify(url));
});
test('the value is the trimmed address, as it is stored', () => {
	assert.deepEqual(plain(ctx.get('  C:\\x\\a.pdf ')), {kind: 'windows', value: 'C:\\x\\a.pdf'});
});
test('the exports and the dialogs use it', () => {
	assert(src.includes('window["AscCommon"].getLocalFileHyperlink = getLocalFileHyperlink'));
	assert(src.includes('window["AscCommon"].isLocalFileHyperlink = isLocalFileHyperlink'));
	// asc_getUrlType itself is unchanged: the upstream tests expect "/home/user/123.txt" to be Invalid there
	const api = fs.readFileSync(path.join(root, 'common/apiBase.js'), 'utf8');
	const body = api.slice(api.indexOf('prototype.asc_getUrlType = function'), api.indexOf('prototype.asc_prepareUrl'));
	assert(body.length > 50 && !body.includes('LocalFileHyperlink'));
});
