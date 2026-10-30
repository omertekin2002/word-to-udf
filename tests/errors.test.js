const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx } = require('./helpers');

const GENERIC = 'Failed to parse DOCX file. Make sure it is a valid Word document.';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * A valid .docx with word/document.xml replaced (or removed, for null)
 */
async function withDocumentXml(documentXml) {
    const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>'));
    if (documentXml === null) {
        zip.remove('word/document.xml');
    } else {
        zip.file('word/document.xml', documentXml);
    }
    return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * A minimal OLE compound file: the 512-byte header, then one directory sector holding
 * a root entry and a stream entry for each name
 */
function compoundFile(streamNames, { trailer = Buffer.alloc(0) } = {}) {
    const header = Buffer.alloc(512);
    Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]).copy(header);
    const entry = (name, type) => {
        const bytes = Buffer.alloc(128);
        bytes.write(name, 0, 'utf16le');
        bytes.writeUInt16LE((name.length + 1) * 2, 64);
        bytes[66] = type;
        return bytes;
    };
    const directory = [entry('Root Entry', 5), ...streamNames.map(name => entry(name, 2))];
    return Buffer.concat([header, ...directory, Buffer.alloc(512 - 128 * directory.length), trailer]);
}

async function assertRejects(input, code) {
    const { DocxParser } = loadConverter();
    await assert.rejects(new DocxParser().parse(input), error => {
        assert.equal(error.code, code);
        if (code === undefined) assert.equal(error.message, GENERIC);
        return true;
    });
}

// The parser logs the underlying error before rethrowing; keep the test output clean
test.beforeEach(t => t.mock.method(console, 'error', () => {}));

test('a truncated word/document.xml is rejected as damaged, not converted to a blank UDF', async () => {
    const docx = await withDocumentXml(`<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Dilek`);
    await assertRejects(docx, 'corrupt');
});

test('a word/document.xml whose root is not w:document is rejected as damaged', async () => {
    await assertRejects(await withDocumentXml(`<w:styles ${W}><w:body/></w:styles>`), 'corrupt');
    await assertRejects(await withDocumentXml('<document><body><p>text</p></body></document>'), 'corrupt');
});

test('a .docx without word/document.xml is rejected as damaged', async () => {
    await assertRejects(await withDocumentXml(null), 'corrupt');
});

test('a compound file with an EncryptionInfo stream is rejected as password-protected', async () => {
    await assertRejects(compoundFile(['EncryptionInfo', 'EncryptedPackage']), 'encrypted');
});

test('any other compound file is rejected as a legacy .doc', async () => {
    await assertRejects(compoundFile(['WordDocument', '1Table']), 'legacyDoc');
});

test('EncryptionInfo in a .doc\'s text, outside a directory entry, does not make it password-protected', async () => {
    const text = Buffer.concat([Buffer.alloc(3), Buffer.from('EncryptionInfo', 'utf16le'), Buffer.alloc(200)]);
    await assertRejects(compoundFile(['WordDocument'], { trailer: text }), 'legacyDoc');
});

test('other unreadable input keeps the generic parse error', async () => {
    await assertRejects(Buffer.from('this is not a zip archive'), undefined);
    await assertRejects(Buffer.from([0xD0, 0xCF, 0x11]), undefined);
    await assertRejects(Buffer.alloc(0), undefined);
});

test('a valid document still converts after a failed one with the same parser', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    await assert.rejects(parser.parse(await withDocumentXml(null)));
    const parsed = await parser.parse(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>'));
    assert.equal(parsed.elements.length, 1);
});
