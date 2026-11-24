const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadConverter, buildDocx, convertBody, textOf } = require('./helpers');

// Embedded content from another program (w:altChunk): reported, never silently dropped, and a
// document with nothing else is refused

const p = runs => `<w:p>${runs}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const altChunk = id => `<w:altChunk r:id="${id}"/>`;
const sdt = inner => `<w:sdt><w:sdtPr><w:alias w:val="Dilekçe"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const MHT = 'MIME-Version: 1.0\r\nContent-Type: text/html\r\n\r\n<html><body><p>Gömülü metin</p></body></html>';
const CHUNKS = { 'afchunk.mht': MHT };

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const bodyParagraphs = (doc, content) => [...doc.querySelectorAll('elements > paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));

test('a chunk between two paragraphs is reported with its part name and place; the paragraphs are converted', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Before')) + altChunk('rId_afchunk') + p(r('After')), { chunks: CHUNKS });
    assert.deepEqual(bodyParagraphs(doc, content), ['Before', 'After']);
    assert.ok(!content.includes('Gömülü'));
    assert.deepEqual(finding(parsed, 'embeddedContent'), {
        code: 'embeddedContent', count: 1, examples: ['afchunk.mht'],
        occurrences: [{ example: 'afchunk.mht', location: { kind: 'paragraph', paragraph: 2 } }]
    });
});

test('a chunk takes one body paragraph place, under the heading before it', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
        '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';
    const heading = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + r('AÇIKLAMALAR') + '</w:p>';
    const { parsed } = await convertBody(heading + altChunk('rId_afchunk') + p(r('Sonra') + '<w:r><w:rPr><w:vanish/></w:rPr><w:t>gizli</w:t></w:r>'),
        { chunks: CHUNKS, styles });
    assert.deepEqual(finding(parsed, 'embeddedContent').occurrences,
        [{ example: 'afchunk.mht', location: { kind: 'paragraph', paragraph: 2, heading: 'AÇIKLAMALAR' } }]);
    // The paragraph after the chunk counts on from it
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences,
        [{ example: 'gizli', location: { kind: 'paragraph', paragraph: 3, heading: 'AÇIKLAMALAR' } }]);
});

test('a chunk in a table cell is reported at that cell', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Before')) + tbl([p(r('Sol')), altChunk('rId_afchunk')]) + p(r('After')), { chunks: CHUNKS });
    assert.ok(content.includes('Sol') && content.includes('Before') && content.includes('After'));
    assert.equal(doc.querySelectorAll('elements > table cell').length, 2);
    assert.deepEqual(finding(parsed, 'embeddedContent'), {
        code: 'embeddedContent', count: 1, examples: ['afchunk.mht'],
        occurrences: [{ example: 'afchunk.mht', location: { kind: 'cell', table: 1, row: 1, cell: 2 } }]
    });
});

test('a chunk in a block-level content control is reported like one in the body', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Before')) + sdt(altChunk('rId_afchunk')) + p(r('After')), { chunks: CHUNKS });
    assert.deepEqual(bodyParagraphs(doc, content), ['Before', 'After']);
    assert.deepEqual(finding(parsed, 'embeddedContent').occurrences,
        [{ example: 'afchunk.mht', location: { kind: 'paragraph', paragraph: 2 } }]);
});

test('each chunk is counted, and one without a relationship is reported without a name', async () => {
    const { parsed } = await convertBody(p(r('Metin')) + altChunk('rId_afchunk') + altChunk('rId_chunk2') + altChunk('rIdMissing'),
        { chunks: { ...CHUNKS, 'chunk2.rtf': '{\\rtf1 Metin}' } });
    assert.deepEqual(finding(parsed, 'embeddedContent'), {
        code: 'embeddedContent', count: 3, examples: ['afchunk.mht', 'chunk2.rtf'],
        occurrences: [
            { example: 'afchunk.mht', location: { kind: 'paragraph', paragraph: 2 } },
            { example: 'chunk2.rtf', location: { kind: 'paragraph', paragraph: 3 } },
            { example: '', location: { kind: 'paragraph', paragraph: 4 } }
        ]
    });
});

test('a chunk in the converted header is noted as header content left out', async () => {
    const { parsed } = await convertBody(p(r('Gövde')), {
        chunks: CHUNKS,
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>',
        parts: [{ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Üst')) + '<w:altChunk r:id="rId_x"/>' }]
    });
    assert.equal(finding(parsed, 'headersFooters').count, 1);
});

test('a document with a chunk and text of its own is converted, not refused', async () => {
    const { content } = await convertBody(altChunk('rId_afchunk') + p(r('Kendi metni')), { chunks: CHUNKS });
    assert.ok(content.includes('Kendi metni'));
});

test('a document without chunks reports no embedded content', async () => {
    const { parsed } = await convertBody(p(r('Metin')));
    assert.equal(finding(parsed, 'embeddedContent'), undefined);
});

test('a document whose body is only chunks, empty paragraphs and a section is refused as embeddedOnly', async t => {
    t.mock.method(console, 'error', () => {});
    const { DocxParser } = loadConverter();
    for (const body of [
        altChunk('rId_afchunk'),
        '<w:p/>' + altChunk('rId_afchunk') + p(r('   ')) + sdt(altChunk('rId_afchunk')),
        tbl([altChunk('rId_afchunk')])
    ]) {
        const docx = await buildDocx(body, { chunks: CHUNKS, sectPr: '<w:pgSz w:w="11906" w:h="16838"/>' });
        await assert.rejects(new DocxParser().parse(docx), error => {
            assert.equal(error.code, 'embeddedOnly');
            return true;
        });
    }
});
