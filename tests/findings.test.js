const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx, convert } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

const p = inner => `<w:p>${inner}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tc = inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const tbl = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    rows.map(cells => `<w:tr>${cells.map(tc).join('')}</w:tr>`).join('') + '</w:tbl>';
const picture = (rId = 'rId_image1') => '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

// Leading bytes of each picture format
const IMAGES = {
    png: Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
    jpeg: Buffer.from([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46]),
    gif: Buffer.from('GIF89a\x01\x00\x01\x00', 'latin1'),
    tiff: Buffer.from([0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00]),
    tiffBigEndian: Buffer.from([0x4D, 0x4D, 0x00, 0x2A, 0x00, 0x00, 0x00, 0x08]),
    wmf: Buffer.from([0xD7, 0xCD, 0xC6, 0x9A, 0x00, 0x00]),
    wmfNoPlaceable: Buffer.from([0x01, 0x00, 0x09, 0x00, 0x00, 0x03]),
    emf: (() => {
        const bytes = Buffer.alloc(88);
        bytes.writeUInt32LE(1, 0);
        bytes.write(' EMF', 40, 'latin1');
        return bytes;
    })(),
    svg: Buffer.from('<?xml version="1.0" encoding="UTF-8"?>\n<!-- logo -->\n<svg xmlns="http://www.w3.org/2000/svg"/>'),
    unknown: Buffer.from('not a picture')
};

const header = (name, id, inner, type = 'header') => ({ name, root: type === 'header' ? 'hdr' : 'ftr', type, id, inner });
const headerRefs = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';

const comments = list => ({
    name: 'comments.xml', root: 'comments', type: 'comments', id: 'rIdComments',
    inner: list.map((text, i) => `<w:comment w:id="${i}" w:author="Yazar" w:initials="Y">${p(r(text))}</w:comment>`).join('')
});

// The parser runs in its own realm; copy its findings into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));

async function parse(docx, parser = new (loadConverter().DocxParser)()) {
    const parsed = await parser.parse(docx);
    return { ...parsed, findings: plain(parsed.findings) };
}

const findingsOf = async (bodyXml, options) => (await parse(await buildDocx(bodyXml, options))).findings;

test('a clean document has no findings', async () => {
    assert.deepEqual(await findingsOf(p(r('Metin'))), []);
});

test('a header and a footer with text give exactly one headersFooters finding, without a location', async () => {
    const findings = await findingsOf(p(r('Gövde')), {
        sectPr: headerRefs,
        parts: [
            header('header1.xml', 'rIdHeader1', p(r('Üst bilgi'))),
            header('footer1.xml', 'rIdFooter1', p(r('Alt bilgi')), 'footer')
        ]
    });
    assert.deepEqual(findings, [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('only a footer with text is enough, also when a section break in the body references it', async () => {
    const sectionBreak = `<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/></w:sectPr></w:pPr>${r('Bölüm 1')}</w:p>`;
    const findings = await findingsOf(sectionBreak + p(r('Bölüm 2')), {
        parts: [header('footer1.xml', 'rIdFooter1', p('<w:r><w:t>1</w:t></w:r>'), 'footer')]
    });
    assert.deepEqual(findings, [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('empty or whitespace-only header and footer parts give no finding', async () => {
    const findings = await findingsOf(p(r('Gövde')), {
        sectPr: headerRefs,
        parts: [
            header('header1.xml', 'rIdHeader1', p('')),
            header('footer1.xml', 'rIdFooter1', p(r('   ') + '<w:r><w:tab/></w:r>'), 'footer')
        ]
    });
    assert.deepEqual(findings, []);
});

test('a header part no section references gives no finding', async () => {
    const findings = await findingsOf(p(r('Gövde')), {
        parts: [header('header1.xml', 'rIdHeader1', p(r('Kullanılmayan üst bilgi')))]
    });
    assert.deepEqual(findings, []);
});

test('two comments give comments with count 2; a comment without text is not counted', async () => {
    const body = p('<w:commentRangeStart w:id="0"/>' + r('Yorumlu metin') + '<w:commentRangeEnd w:id="0"/>' +
        '<w:r><w:commentReference w:id="0"/></w:r>');
    const findings = await findingsOf(body, { parts: [comments(['Birinci yorum', 'İkinci yorum', '  '])] });
    assert.deepEqual(findings, [{ code: 'comments', count: 2, examples: [] }]);
});

test('the comments part is found through its relationship, and at word/comments.xml without one', async () => {
    const renamed = { ...comments(['Yorum']), name: 'notlar/yorumlar.xml' };
    assert.deepEqual(await findingsOf(p(r('Metin')), { parts: [renamed] }),
        [{ code: 'comments', count: 1, examples: [] }]);

    const zip = await JSZip.loadAsync(await buildDocx(p(r('Metin')), { parts: [comments(['Yorum'])] }));
    const rels = await zip.file('word/_rels/document.xml.rels').async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace(/<Relationship Id="rIdComments"[^>]*\/>/, ''));
    const { findings } = await parse(await zip.generateAsync({ type: 'nodebuffer' }));
    assert.deepEqual(findings, [{ code: 'comments', count: 1, examples: [] }]);
});

test('a table nested in a cell gives nestedTables with the nested table\'s text, and converts as before', async () => {
    const nested = tbl([[p('') + p(r('İç tablo metni')), p(r('İkinci hücre'))]]);
    const body = tbl([[p(r('Dış hücre')) + nested, p(r('Yan hücre'))]]);
    const { parsed, content } = await convert(await buildDocx(body));

    assert.deepEqual(plain(parsed.findings), [{ code: 'nestedTables', count: 1, examples: ['İç tablo metni'] }]);
    // Flattened into the containing cell in reading order, as before
    const cellTexts = plain(parsed.elements[0].rows[0].cells).map(cell =>
        cell.paragraphs.map(para => para.runs.map(run => run.text || '').join('')));
    assert.deepEqual(cellTexts, [['Dış hücre', '', 'İç tablo metni', 'İkinci hücre'], ['Yan hücre']]);
    const order = ['Dış hücre', 'İç tablo metni', 'İkinci hücre', 'Yan hücre'].map(text => content.indexOf(text));
    assert.ok(order.every((at, i) => at >= 0 && (i === 0 || at > order[i - 1])), JSON.stringify(content));
});

test('nested table examples are cut to about 40 characters, and a table nested two deep counts at each level', async () => {
    const long = 'Bu iç tablonun ilk paragrafı kırk karakterden uzundur ve kesilir';
    const deepest = tbl([[p(r('En içteki'))]]);
    const middle = tbl([[p(r(long)) + deepest]]);
    const findings = await findingsOf(tbl([[p(r('Dış')) + middle]]));

    assert.equal(findings.length, 1);
    assert.equal(findings[0].code, 'nestedTables');
    assert.equal(findings[0].count, 2);
    assert.deepEqual(findings[0].examples, ['En içteki', 'Bu iç tablonun ilk paragrafı kırk karakt…']);
    assert.ok(findings[0].examples[1].length <= 41);
});

test('examples are capped at 3 and not repeated, while the count includes every occurrence', async () => {
    const cellWith = text => tbl([[p(r('Dış')) + tbl([[p(r(text))]])]]);
    const findings = await findingsOf(['A', 'B', 'A', 'C', 'D'].map(cellWith).join(''));
    assert.deepEqual(findings, [{ code: 'nestedTables', count: 5, examples: ['A', 'B', 'C'] }]);
});

test('a drawing whose relationship is missing gives missingImage with its paragraph text', async () => {
    const findings = await findingsOf(p(r('Ekteki imza: ') + picture('rId_yok')));
    assert.deepEqual(findings, [{ code: 'missingImage', count: 1, examples: ['Ekteki imza:'] }]);
});

test('a drawing whose media part is missing gives missingImage with the nearest paragraph text, the earlier one on a tie', async () => {
    const zip = await JSZip.loadAsync(await buildDocx(
        p('') + p(r('Önceki paragraf')) + p(picture()) + p('') + p(r('Sonraki paragraf')),
        { media: { 'image1.png': IMAGES.png } }));
    zip.remove('word/media/image1.png');
    const { findings } = await parse(await zip.generateAsync({ type: 'nodebuffer' }));
    assert.deepEqual(findings, [{ code: 'missingImage', count: 1, examples: ['Önceki paragraf'] }]);
});

test('a drawing with only a text box, and no picture, gives no finding', async () => {
    const textBox = '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        `<wps:wsp><wps:txbx><w:txbxContent>${p(r('Kutu'))}</w:txbxContent></wps:txbx></wps:wsp>` +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    assert.deepEqual(await findingsOf(p(r('Çapa') + textBox)), []);
});

for (const [name, format] of [
    ['emf', 'EMF'], ['wmf', 'WMF'], ['wmfNoPlaceable', 'WMF'], ['gif', 'GIF'], ['tiff', 'TIFF'],
    ['tiffBigEndian', 'TIFF'], ['svg', 'SVG'], ['unknown', 'unknown']
]) {
    test(`a ${format} picture (${name}) gives unsupportedImage naming ${format}`, async () => {
        // Named .png on purpose: the format comes from the bytes, not the extension
        const findings = await findingsOf(p(picture()), { media: { 'image1.png': IMAGES[name] } });
        assert.deepEqual(findings, [{ code: 'unsupportedImage', count: 1, examples: [format] }]);
    });
}

test('PNG and JPEG pictures give no finding, whatever their extension', async () => {
    const findings = await findingsOf(p(picture('rId_image1') + picture('rId_image2')),
        { media: { 'image1.gif': IMAGES.png, 'image2.emf': IMAGES.jpeg } });
    assert.deepEqual(findings, []);
});

test('findings come in a fixed order, each code once', async () => {
    const body = tbl([[p(r('Dış')) + tbl([[p(r('İç'))]])]]) + p(r('Resim') + picture('rId_yok')) +
        p(picture('rId_image1') + picture('rId_image2'));
    const findings = await findingsOf(body, {
        sectPr: headerRefs,
        media: { 'image1.emf': IMAGES.emf, 'image2.gif': IMAGES.gif },
        parts: [header('header1.xml', 'rIdHeader1', p(r('Üst'))), header('footer1.xml', 'rIdFooter1', p(''), 'footer'),
            comments(['Yorum'])]
    });
    assert.deepEqual(findings, [
        { code: 'headersFooters', count: 1, examples: [] },
        { code: 'comments', count: 1, examples: [] },
        { code: 'nestedTables', count: 1, examples: ['İç'] },
        { code: 'missingImage', count: 1, examples: ['Resim'] },
        { code: 'unsupportedImage', count: 2, examples: ['EMF', 'GIF'] }
    ]);
});

test('the same parser gives no findings for a clean document after one with findings', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const first = await parse(await buildDocx(p(picture('rId_yok')), { parts: [comments(['Yorum'])] }), parser);
    assert.equal(first.findings.length, 2);

    const second = await parse(await buildDocx(p(r('Temiz'))), parser);
    assert.deepEqual(second.findings, []);
});

test('the generator ignores findings: the UDF is the same with or without them', async () => {
    const { DocxParser, UdfGenerator } = loadConverter();
    const docx = await buildDocx(tbl([[p(r('Dış')) + tbl([[p(r('İç'))]])]]) + p(picture()),
        { media: { 'image1.png': IMAGES.gif } });
    const parsed = await new DocxParser().parse(docx);
    assert.equal(parsed.findings.length, 2);
    const udfXml = async document => (await JSZip.loadAsync(
        await (await new UdfGenerator().generate(document)).arrayBuffer())).file('content.xml').async('string');
    assert.equal(await udfXml(parsed), await udfXml({ ...parsed, findings: [] }));
});

// None of the petition fixtures has headers, comments, nested tables or pictures
const EXPECTED_FIXTURE_FINDINGS = {
    'dava dilekçesi': [],
    'icra/itiraz dilekçesi': [],
    'ekli tablo dilekçesi': [],
    'cevap dilekçesi': []
};

test('every petition fixture has an expected findings entry', () => {
    assert.deepEqual(petitions.map(petition => petition.name).sort(), Object.keys(EXPECTED_FIXTURE_FINDINGS).sort());
});

for (const petition of petitions) {
    test(`the ${petition.name} fixture gives exactly its expected findings`, async () => {
        const { findings } = await parse(await petition.build());
        assert.deepEqual(findings, EXPECTED_FIXTURE_FINDINGS[petition.name]);
    });
}
