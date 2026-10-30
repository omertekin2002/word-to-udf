const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx, compoundFile } = require('./helpers');

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

test('a password-protected .docx is recognised when its directory spans several sectors', async () => {
    // The layout Word writes: the \x06DataSpaces storage first, so EncryptionInfo is past the first sector
    const dataSpaces = { name: '\x06DataSpaces', children: [
        'Version', 'DataSpaceMap',
        { name: 'DataSpaceInfo', children: ['StrongEncryptionDataSpace'] },
        { name: 'TransformInfo', children: [{ name: 'StrongEncryptionTransform', children: ['\x06Primary'] }] }
    ] };
    await assertRejects(compoundFile([dataSpaces, 'EncryptionInfo', 'EncryptedPackage']), 'encrypted');
});

test('a .doc with an embedded password-protected object in a storage is rejected as a legacy .doc', async () => {
    const objectPool = { name: 'ObjectPool', children: [{ name: '_1234567890', children: ['EncryptionInfo', 'EncryptedPackage'] }] };
    await assertRejects(compoundFile(['WordDocument', '1Table', objectPool]), 'legacyDoc');
});

test('a .doc carrying an embedded password-protected file\'s bytes is rejected as a legacy .doc', async () => {
    // The embedded compound file starts on a sector boundary, so its directory entries line up too
    const embedded = compoundFile(['EncryptionInfo', 'EncryptedPackage']);
    await assertRejects(compoundFile(['WordDocument', '1Table'], { trailer: embedded }), 'legacyDoc');
});

test('any other compound file is rejected as a legacy .doc', async () => {
    await assertRejects(compoundFile(['WordDocument', '1Table']), 'legacyDoc');
});

test('EncryptionInfo in a .doc\'s text, outside a directory entry, does not make it password-protected', async () => {
    const text = Buffer.concat([Buffer.alloc(3), Buffer.from('EncryptionInfo', 'utf16le'), Buffer.alloc(200)]);
    await assertRejects(compoundFile(['WordDocument'], { trailer: text }), 'legacyDoc');
});

test('a Strict Open XML document is rejected as unsupported, not damaged', async () => {
    const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>'));
    const documentXml = await zip.file('word/document.xml').async('string');
    zip.file('word/document.xml', documentXml.replace('http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'http://purl.oclc.org/ooxml/wordprocessingml/main'));
    await assertRejects(await zip.generateAsync({ type: 'nodebuffer' }), 'strict');
});

test('a main document part named in _rels/.rels converts, wherever it is', async () => {
    const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>',
        { media: { 'image1.png': Buffer.from('png') } }));
    for (const [from, to] of [['word/document.xml', 'word/document2.xml'],
        ['word/_rels/document.xml.rels', 'word/_rels/document2.xml.rels']]) {
        zip.file(to, await zip.file(from).async('nodebuffer'));
        zip.remove(from);
    }
    const rels = await zip.file('_rels/.rels').async('string');
    zip.file('_rels/.rels', rels.replace('Target="word/document.xml"', 'Target="/word/document2.xml"'));

    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const parsed = await parser.parse(await zip.generateAsync({ type: 'nodebuffer' }));
    assert.equal(parsed.elements.length, 1);
    assert.equal(parser.relationships.rId_image1.target, 'media/image1.png');
});

test('a main document part outside word/ reads the supporting parts its relationships name', async () => {
    const picture = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:blipFill><a:blip r:embed="rId_image1"/></pic:blipFill></pic:pic>' +
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const body = '<w:p><w:pPr><w:pStyle w:val="Heading"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
        `<w:p>${picture}</w:p>`;
    const zip = await JSZip.loadAsync(await buildDocx(body, {
        styles: '<w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:b/></w:rPr></w:style>',
        numbering: '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/>' +
            '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
        footnotes: { 1: '<w:p><w:r><w:t>Source text</w:t></w:r></w:p>' },
        settings: '<w:gutterAtTop/>',
        sectPr: '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="567"/>',
        media: { 'image1.png': Buffer.from('png') }
    }));

    // Every word/ part moves to docs/, but the picture to shared/, named with a relative
    // target that leaves the document's folder; numbering is named with an absolute target
    for (const from of Object.keys(zip.files).filter(name => name.startsWith('word/') && !zip.files[name].dir)) {
        const to = from === 'word/media/image1.png' ? 'shared/image1.png' : from.replace(/^word\//, 'docs/');
        zip.file(to, await zip.file(from).async('nodebuffer'));
        zip.remove(from);
    }
    zip.remove('word');
    const documentRels = await zip.file('docs/_rels/document.xml.rels').async('string');
    zip.file('docs/_rels/document.xml.rels', documentRels
        .replace('Target="numbering.xml"', 'Target="/docs/numbering.xml"')
        .replace('Target="media/image1.png"', 'Target="../shared/image1.png"'));
    const rels = await zip.file('_rels/.rels').async('string');
    zip.file('_rels/.rels', rels.replace('Target="word/document.xml"', 'Target="docs/document.xml"'));
    const contentTypes = await zip.file('[Content_Types].xml').async('string');
    zip.file('[Content_Types].xml', contentTypes.replace(/"\/word\//g, '"/docs/'));

    const parsed = await new (loadConverter().DocxParser)().parse(await zip.generateAsync({ type: 'nodebuffer' }));
    const [claim, figure] = parsed.elements;
    assert.equal(claim.runs[0].bold, true, 'the paragraph style from docs/styles.xml');
    assert.equal(claim.numbering && claim.numbering.format, 'decimal', 'the list from docs/numbering.xml');
    assert.equal(claim.runs[1].content, 'Source text', 'the footnote from docs/footnotes.xml');
    assert.equal(figure.runs[0].data, Buffer.from('png').toString('base64'), 'the picture from shared/image1.png');
    assert.equal(parsed.pageSetup.topMargin, 72 + 567 / 20, 'w:gutterAtTop from docs/settings.xml');
});

for (const encoding of ['utf16le', 'utf16be']) {
    test(`a ${encoding} word/document.xml converts`, async () => {
        const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>'));
        const documentXml = (await zip.file('word/document.xml').async('string')).replace('UTF-8', 'UTF-16');
        const bytes = Buffer.from('\uFEFF' + documentXml, 'utf16le');
        if (encoding === 'utf16be') bytes.swap16();
        zip.file('word/document.xml', bytes);

        const parsed = await new (loadConverter().DocxParser)().parse(await zip.generateAsync({ type: 'nodebuffer' }));
        assert.equal(parsed.elements.length, 1);
        assert.match(JSON.stringify(parsed.elements), /Dilekçe metni/);
    });
}

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
