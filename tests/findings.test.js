const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx, convert, findingSummaries } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

const p = inner => `<w:p>${inner}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tc = inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const tbl = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    rows.map(cells => `<w:tr>${cells.map(tc).join('')}</w:tr>`).join('') + '</w:tbl>';
// Content in a cell of a table in a cell of another table. In a body table's cell, a table in it
// is the 4th level of tables, past the depth limit: it is flattened into its cell and reported
const inTwoTables = inner => tbl([[tbl([[inner]])]]);
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

// Each finding's code, count and examples; the located occurrences are checked on their own below
async function parse(docx, parser = new (loadConverter().DocxParser)()) {
    const parsed = await parser.parse(docx);
    return { ...parsed, findings: findingSummaries(parsed.findings) };
}

const findingsOf = async (bodyXml, options) => (await parse(await buildDocx(bodyXml, options))).findings;

test('a clean document has no findings', async () => {
    assert.deepEqual(await findingsOf(p(r('Metin'))), []);
});

test('a default header and footer with plain text are converted and give no finding', async () => {
    const findings = await findingsOf(p(r('Gövde')), {
        sectPr: headerRefs,
        parts: [
            header('header1.xml', 'rIdHeader1', p(r('Üst bilgi'))),
            header('footer1.xml', 'rIdFooter1', p(r('Alt bilgi')), 'footer')
        ]
    });
    assert.deepEqual(findings, []);
});

test('first-page and even-page headers and footers with text give exactly one headersFooters finding, without a location', async () => {
    const findings = await findingsOf(p(r('Gövde')), {
        sectPr: '<w:headerReference w:type="first" r:id="rIdHeader1"/><w:footerReference w:type="even" r:id="rIdFooter1"/>',
        parts: [
            header('header1.xml', 'rIdHeader1', p(r('İlk sayfa üst bilgisi'))),
            header('footer1.xml', 'rIdFooter1', p(r('Çift sayfa alt bilgisi')), 'footer')
        ]
    });
    assert.deepEqual(findings, [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('an earlier section\'s footer with text gives the finding when the final section has a different one', async () => {
    const sectionBreak = `<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/></w:sectPr></w:pPr>${r('Bölüm 1')}</w:p>`;
    const findings = await findingsOf(sectionBreak + p(r('Bölüm 2')), {
        sectPr: '<w:footerReference w:type="default" r:id="rIdFooter2"/>',
        parts: [
            header('footer1.xml', 'rIdFooter1', p(r('Birinci bölüm')), 'footer'),
            header('footer2.xml', 'rIdFooter2', p(r('Son bölüm')), 'footer')
        ]
    });
    assert.deepEqual(findings, [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('an earlier section\'s footer the final section inherits is converted, without a finding', async () => {
    const sectionBreak = `<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/></w:sectPr></w:pPr>${r('Bölüm 1')}</w:p>`;
    const findings = await findingsOf(sectionBreak + p(r('Bölüm 2')), {
        parts: [header('footer1.xml', 'rIdFooter1', p(r('Ortak alt bilgi')), 'footer')]
    });
    assert.deepEqual(findings, []);
});

test('a page number field, a table or a text box in the converted header or footer gives the finding', async () => {
    const pageField = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const cases = {
        'PAGE field': ['footer', p(r('Sayfa ') + pageField)],
        'NUMPAGES simple field': ['footer', p(r('Toplam ') + '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>')],
        table: ['header', tbl([[p(r('Hücre')), p('')]]) + p(r('Üst bilgi'))],
        'text box': ['header', p(r('Üst bilgi') + '<w:r><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox>' +
            `<w:txbxContent>${p(r('Kutu'))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`)]
    };
    for (const [name, [kind, inner]] of Object.entries(cases)) {
        const findings = await findingsOf(p(r('Gövde')), {
            sectPr: headerRefs,
            parts: [kind === 'header' ? header('header1.xml', 'rIdHeader1', inner) : header('footer1.xml', 'rIdFooter1', inner, 'footer')]
        });
        assert.deepEqual(findings, [{ code: 'headersFooters', count: 1, examples: [] }], name);
    }
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

test('tables nested within the depth limit are written as tables and give no nestedTables', async () => {
    assert.deepEqual(await findingsOf(tbl([[p(r('Dış')) + tbl([[p(r('İç'))]]) + p('')]])), []);
    assert.deepEqual(await findingsOf(tbl([[p(r('Dış')) + tbl([[p(r('Orta')) + tbl([[p(r('İç'))]])]]) + p('')]])), []);
});

test('a table nested past the depth limit gives nestedTables with its text, and is flattened into its cell', async () => {
    const nested = tbl([[p('') + p(r('İç tablo metni')), p(r('İkinci hücre'))]]);
    const third = tbl([[p(r('Dış hücre')) + nested, p(r('Yan hücre'))]]);
    const body = tbl([[tbl([[third]])]]);
    const { parsed, content } = await convert(await buildDocx(body));

    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'nestedTables', count: 1, examples: ['İç tablo metni'] }]);
    // Flattened into the containing cell of the 3rd table in reading order
    const deepest = parsed.elements[0].rows[0].cells[0].blocks[0].rows[0].cells[0].blocks[0];
    const cellTexts = plain(deepest.rows[0].cells).map(cell =>
        cell.blocks.map(block => block.type === 'paragraph' ? block.runs.map(run => run.text || '').join('') : block.type));
    assert.deepEqual(cellTexts, [['Dış hücre', '', 'İç tablo metni', 'İkinci hücre'], ['Yan hücre']]);
    const order = ['Dış hücre', 'İç tablo metni', 'İkinci hücre', 'Yan hücre'].map(text => content.indexOf(text));
    assert.ok(order.every((at, i) => at >= 0 && (i === 0 || at > order[i - 1])), JSON.stringify(content));
});

test('nested table examples are cut to about 40 characters, and a table nested two past the limit counts at each level', async () => {
    const long = 'Bu iç tablonun ilk paragrafı kırk karakterden uzundur ve kesilir';
    const deepest = tbl([[p(r('En içteki'))]]);
    const middle = tbl([[p(r(long)) + deepest]]);
    const findings = await findingsOf(tbl([[p(r('Dış')) + inTwoTables(middle)]]));

    assert.equal(findings.length, 1);
    assert.equal(findings[0].code, 'nestedTables');
    assert.equal(findings[0].count, 2);
    assert.deepEqual(findings[0].examples, ['En içteki', 'Bu iç tablonun ilk paragrafı kırk karakt…']);
    assert.ok(findings[0].examples[1].length <= 41);
});

test('examples are capped at 3 and not repeated, while the count includes every occurrence', async () => {
    const cellWith = text => tbl([[p(r('Dış')) + inTwoTables(tbl([[p(r(text))]]))]]);
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
    const body = tbl([[p(r('Dış')) + inTwoTables(tbl([[p(r('İç'))]]))]]) + p(r('Resim') + picture('rId_yok')) +
        p(picture('rId_image1') + picture('rId_image2'));
    const findings = await findingsOf(body, {
        sectPr: '<w:headerReference w:type="first" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>',
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
    const docx = await buildDocx(tbl([[p(r('Dış')) + inTwoTables(tbl([[p(r('İç'))]]))]]) + p(picture()),
        { media: { 'image1.png': IMAGES.gif } });
    const parsed = await new DocxParser().parse(docx);
    assert.equal(parsed.findings.length, 2);
    const udfXml = async document => (await JSZip.loadAsync(
        await (await new UdfGenerator().generate(document)).arrayBuffer())).file('content.xml').async('string');
    assert.equal(await udfXml(parsed), await udfXml({ ...parsed, findings: [] }));
});

// None of the petition fixtures has headers, comments, nested tables, pictures, endnotes, exact or
// at-least line heights, or bar stops and underscore or heavy leaders. Only the dava dilekçesi has footnotes,
// and Wingdings check boxes in its EKLER list.
const EXPECTED_FIXTURE_FINDINGS = {
    'dava dilekçesi': [{
        code: 'notesMoved',
        count: 2,
        examples: ['Ek-1: Hizmet sözleşmesinin örneği.', 'Ek-2: Örnek Noterliğinin 00.00.2026 tari…']
    }, {
        code: 'symbols',
        count: 3,
        examples: ['Wingdings F0FE → ☑', 'Wingdings F0A8 → □']
    }],
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

// Footnotes, endnotes, line heights and tab stops the conversion moves or loses

const footnoteRef = id => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
const endnoteRef = id => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;
const sectionBreak = (inner = '') => `<w:p><w:pPr><w:sectPr>${inner}</w:sectPr></w:pPr>${r('Bölüm sonu')}</w:p>`;
const endnotePr = inner => `<w:endnotePr>${inner}</w:endnotePr>`;
const pos = value => `<w:pos w:val="${value}"/>`;
const numRestart = value => `<w:numRestart w:val="${value}"/>`;
const ENDNOTES = { 1: p(r('Birinci son not')), 2: p(r('İkinci son not')) };

test('each footnote reference gives notesMoved with the note text, also in a table cell', async () => {
    const long = 'Bu dipnotun metni kırk karakterden uzundur ve kısaltılır';
    const findings = await findingsOf(
        p(r('Birinci') + footnoteRef(1)) + tbl([[p(r('Hücre') + footnoteRef(2)), p(footnoteRef(1))]]) + p(footnoteRef(3)),
        { footnotes: { 1: p(r('Ek-1: Sözleşme.')), 2: p(r(long)), 3: p('') } });
    assert.deepEqual(findings, [{
        code: 'notesMoved',
        count: 4,
        examples: ['Ek-1: Sözleşme.', 'Bu dipnotun metni kırk karakterden uzund…']
    }]);
});

test('a reference to a missing footnote, and footnotes no paragraph references, give no finding', async () => {
    const findings = await findingsOf(p(r('Metin') + footnoteRef(9)), { footnotes: { 1: p(r('Kullanılmayan')) } });
    assert.deepEqual(findings, []);
});

test('endnotes at the end of the document with continuous numbering give no finding', async () => {
    const body = p(r('Bir') + endnoteRef(1)) + p(r('İki') + endnoteRef(2));
    assert.deepEqual(await findingsOf(body, { endnotes: ENDNOTES }), [], 'Word defaults');
    assert.deepEqual(await findingsOf(sectionBreak() + body, {
        endnotes: ENDNOTES,
        settings: endnotePr(pos('docEnd') + numRestart('continuous'))
    }), [], 'stated defaults, two sections');
});

test('a two-section document with endnotes at section end gives endnotesMoved per endnote reference', async () => {
    const body = sectionBreak() + p(r('Bir') + endnoteRef(1)) + p(r('İki') + endnoteRef(2));
    const expected = [{ code: 'endnotesMoved', count: 2, examples: ['Birinci son not', 'İkinci son not'] }];
    assert.deepEqual(await findingsOf(body, { endnotes: ENDNOTES, settings: endnotePr(pos('sectEnd')) }), expected);
});

test('endnote placement is the document\'s: a w:pos in a section\'s w:endnotePr is ignored', async () => {
    // ECMA-376 Part 1, 17.11.22: "If this element is present at the section level, then it shall be ignored"
    const expected = [{ code: 'endnotesMoved', count: 2, examples: ['Birinci son not', 'İkinci son not'] }];
    const refs = p(r('Bir') + endnoteRef(1)) + p(r('İki') + endnoteRef(2));

    // Every section says docEnd, but the settings part places endnotes at section end: they are moved
    assert.deepEqual(await findingsOf(sectionBreak(endnotePr(pos('docEnd'))) + refs, {
        endnotes: ENDNOTES,
        settings: endnotePr(pos('sectEnd')),
        sectPr: endnotePr(pos('docEnd'))
    }), expected, 'a section-level docEnd does not hide the document\'s sectEnd');

    // Only the sections say sectEnd: Word keeps the endnotes at the end of the document, as the conversion does
    assert.deepEqual(await findingsOf(sectionBreak(endnotePr(pos('sectEnd'))) + refs, {
        endnotes: ENDNOTES,
        sectPr: endnotePr(pos('sectEnd'))
    }), [], 'a section-level sectEnd, no settings part');
    assert.deepEqual(await findingsOf(sectionBreak(endnotePr(pos('sectEnd'))) + refs, {
        endnotes: ENDNOTES,
        settings: endnotePr(pos('docEnd')),
        sectPr: endnotePr(pos('sectEnd'))
    }), [], 'a section-level sectEnd over the document\'s docEnd');

    // The numbering restart next to an ignored w:pos is still the section's own
    assert.deepEqual(await findingsOf(sectionBreak(endnotePr(pos('sectEnd') + numRestart('eachSect'))) + refs,
        { endnotes: ENDNOTES }), expected, 'w:numRestart beside an ignored w:pos');
});

test('section-end endnotes in a document with one section give no finding', async () => {
    const body = p(r('Bir') + endnoteRef(1));
    assert.deepEqual(await findingsOf(body, { endnotes: ENDNOTES, settings: endnotePr(pos('sectEnd')) }), []);
});

test('restarting endnote numbering gives endnotesMoved: at each section with two sections, at each page always', async () => {
    const twoSections = sectionBreak(endnotePr(numRestart('eachSect'))) + p(r('Bir') + endnoteRef(1));
    assert.deepEqual(await findingsOf(twoSections, { endnotes: ENDNOTES }),
        [{ code: 'endnotesMoved', count: 1, examples: ['Birinci son not'] }], 'a section break restarts');
    assert.deepEqual(await findingsOf(sectionBreak() + p(r('Bir') + endnoteRef(1)),
        { endnotes: ENDNOTES, settings: endnotePr(numRestart('eachSect')) }),
    [{ code: 'endnotesMoved', count: 1, examples: ['Birinci son not'] }], 'the settings part restarts');
    assert.deepEqual(await findingsOf(p(r('Bir') + endnoteRef(1)),
        { endnotes: ENDNOTES, settings: endnotePr(numRestart('eachSect')) }), [], 'one section: nothing restarts');
    assert.deepEqual(await findingsOf(p(r('Bir') + endnoteRef(1)),
        { endnotes: ENDNOTES, sectPr: endnotePr(numRestart('eachPage')) }),
    [{ code: 'endnotesMoved', count: 1, examples: ['Birinci son not'] }], 'each page');
});

test('a section that sets continuous numbering itself overrides a restart in the settings part', async () => {
    const body = sectionBreak(endnotePr(numRestart('continuous'))) + p(r('Bir') + endnoteRef(1));
    assert.deepEqual(await findingsOf(body, {
        endnotes: ENDNOTES,
        settings: endnotePr(numRestart('eachSect')),
        sectPr: endnotePr(numRestart('continuous'))
    }), []);
});

test('moved endnotes without a reference in the body give no finding', async () => {
    const findings = await findingsOf(sectionBreak() + p(r('Metin')),
        { endnotes: ENDNOTES, settings: endnotePr(pos('sectEnd')) });
    assert.deepEqual(findings, []);
});

const spacing = (rule, line = 360) => `<w:spacing w:line="${line}"${rule ? ` w:lineRule="${rule}"` : ''}/>`;
const withPPr = (pPr, text) => `<w:p><w:pPr>${pPr}</w:pPr>${r(text)}</w:p>`;
const paraStyle = (id, pPr, basedOn) => `<w:style w:type="paragraph" w:styleId="${id}">` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : '') + `<w:pPr>${pPr}</w:pPr></w:style>`;
const docDefaults = pPr => `<w:docDefaults><w:pPrDefault><w:pPr>${pPr}</w:pPr></w:pPrDefault></w:docDefaults>`;

test('exact and at-least line heights set on the paragraph give lineSpacing with the paragraph text', async () => {
    const findings = await findingsOf(withPPr(spacing('exact'), 'Tam satır') + withPPr(spacing('atLeast'), 'En az satır') +
        withPPr(spacing('auto'), 'Otomatik') + withPPr(spacing(''), 'Kuralsız') + p(r('Ayarsız')));
    assert.deepEqual(findings, [{ code: 'lineSpacing', count: 2, examples: ['Tam satır', 'En az satır'] }]);
});

test('a line rule from the paragraph style chain or w:docDefaults counts, unless the paragraph sets auto', async () => {
    const styles = paraStyle('Base', spacing('exact')) + paraStyle('Child', '', 'Base');
    const findings = await findingsOf(
        withPPr('<w:pStyle w:val="Child"/>', 'Stilden tam') +
        withPPr('<w:pStyle w:val="Child"/>' + spacing('auto', 240), 'Paragrafta otomatik') +
        withPPr('<w:pStyle w:val="Child"/><w:spacing w:after="120"/>', 'Yalnız sonrası'), { styles });
    assert.deepEqual(findings, [{ code: 'lineSpacing', count: 2, examples: ['Stilden tam', 'Yalnız sonrası'] }]);

    const fromDefaults = await findingsOf(p(r('Varsayılan')) + withPPr(spacing('auto'), 'Geçersiz kılınan'),
        { styles: docDefaults(spacing('atLeast')) });
    assert.deepEqual(fromDefaults, [{ code: 'lineSpacing', count: 1, examples: ['Varsayılan'] }]);
});

test('an exact line height in a table cell gives lineSpacing', async () => {
    const findings = await findingsOf(tbl([[withPPr(spacing('exact'), 'Hücrede tam'), p(r('Düz'))]]));
    assert.deepEqual(findings, [{ code: 'lineSpacing', count: 1, examples: ['Hücrede tam'] }]);
});

const tabsOf = (...stops) => `<w:tabs>${stops.join('')}</w:tabs>`;
const stop = (val, position, leader) => `<w:tab w:val="${val}" w:pos="${position}"${leader ? ` w:leader="${leader}"` : ''}/>`;

test('a bar stop or an underscore or heavy leader gives tabStops with the paragraph text', async () => {
    const findings = await findingsOf(
        withPPr(tabsOf(stop('bar', 720), stop('left', 1440)), 'Çubuk') +
        withPPr(tabsOf(stop('right', 9000, 'underscore')), 'Alt çizgi') +
        withPPr(tabsOf(stop('left', 1440, 'heavy')), 'Kalın') +
        withPPr(tabsOf(stop('right', 9000, 'dot'), stop('left', 720, 'middleDot'), stop('center', 4000, 'hyphen'),
            stop('left', 1440, 'none'), stop('left', 2000)), 'Korunan'));
    assert.deepEqual(findings, [{ code: 'tabStops', count: 3, examples: ['Çubuk', 'Alt çizgi', 'Kalın'] }]);
});

test('a lost stop from a style or w:docDefaults counts, unless the paragraph clears or replaces it', async () => {
    const styles = docDefaults(tabsOf(stop('left', 360, 'heavy'))) + paraStyle('Bar', tabsOf(stop('bar', 720)));
    const findings = await findingsOf(
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('clear', 360)), 'Stilden çubuk') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('clear', 720), stop('clear', 360)), 'Temizlenen') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('left', 720, 'dot'), stop('left', 360)), 'Değiştirilen') +
        p(r('Varsayılandan kalın')), { styles });
    assert.deepEqual(findings, [{ code: 'tabStops', count: 2, examples: ['Stilden çubuk', 'Varsayılandan kalın'] }]);
});

test('a clear or a stop one twip from a lost stop does not remove it, so it still counts', async () => {
    // 1441 and 1442 twips both round to 72.1 points, but they are different stops
    const styles = paraStyle('Heavy', tabsOf(stop('left', 1442, 'heavy'))) + paraStyle('Bar', tabsOf(stop('bar', 1442)));
    const findings = await findingsOf(
        withPPr('<w:pStyle w:val="Heavy"/>' + tabsOf(stop('clear', 1441)), 'Yanından temizlenen') +
        withPPr('<w:pStyle w:val="Heavy"/>' + tabsOf(stop('clear', 1442)), 'Tam yerinden temizlenen') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('left', 1441)), 'Yanına eklenen') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('left', 1442)), 'Tam yerinde değiştirilen'), { styles });
    assert.deepEqual(findings, [{ code: 'tabStops', count: 2, examples: ['Yanından temizlenen', 'Yanına eklenen'] }]);
});

test('a clear or a stop at the position of a lost stop removes it when written in another unit', async () => {
    // 432 twips and 0.3in are the same position, but convert to 21.6 and 21.599999999999998 points
    const styles = paraStyle('Heavy', tabsOf(stop('left', 432, 'heavy'))) + paraStyle('Bar', tabsOf(stop('bar', '127mm'))) +
        paraStyle('Plain', tabsOf(stop('left', '0.3in')));
    const findings = await findingsOf(
        withPPr('<w:pStyle w:val="Heavy"/>', 'Kalın kalan') +
        withPPr('<w:pStyle w:val="Heavy"/>' + tabsOf(stop('clear', '0.3in')), 'İnç ile temizlenen') +
        withPPr('<w:pStyle w:val="Heavy"/>' + tabsOf(stop('left', '0.3in', 'dot')), 'İnç ile değiştirilen') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('clear', 7200)), 'Twip ile temizlenen') +
        withPPr('<w:pStyle w:val="Bar"/>' + tabsOf(stop('right', 7200)), 'Twip ile değiştirilen') +
        withPPr('<w:pStyle w:val="Plain"/>' + tabsOf(stop('bar', 432)), 'Twip ile çubuk'), { styles });
    assert.deepEqual(findings, [{ code: 'tabStops', count: 2, examples: ['Kalın kalan', 'Twip ile çubuk'] }]);
});

test('a bar stop in a table cell gives tabStops', async () => {
    const findings = await findingsOf(tbl([[withPPr(tabsOf(stop('bar', 720)), 'Hücrede çubuk'), p(r('Düz'))]]));
    assert.deepEqual(findings, [{ code: 'tabStops', count: 1, examples: ['Hücrede çubuk'] }]);
});

// A one-column table; each cell is [w:vMerge value, content]
const mergeColumn = cells => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' + cells.map(([merge, inner]) =>
    `<w:tr><w:tc><w:tcPr><w:vMerge w:val="${merge}"/></w:tcPr>${inner}</w:tc></w:tr>`).join('') + '</w:tbl>';
// A paragraph with an exact line height and a bar stop: one lineSpacing and one tabStops finding when converted
const lossy = text => withPPr(spacing('exact') + tabsOf(stop('bar', 720)), text);
const lossyFindings = text => [
    { code: 'lineSpacing', count: 1, examples: [text] },
    { code: 'tabStops', count: 1, examples: [text] }
];

test('a cell joined to the vertical merge above it is written empty, and its paragraphs give no findings', async () => {
    const { parsed, doc, content } = await convert(await buildDocx(mergeColumn([
        ['restart', lossy('Birleşik')],
        ['continue', lossy('')],
        ['continue', lossy('Yazılmayan')]
    ])));

    // Only the merge's first cell is converted, so only its paragraph counts
    assert.deepEqual(findingSummaries(parsed.findings), lossyFindings('Birleşik'));
    assert.equal(content.trim(), 'Birleşik');
    const cells = [...doc.querySelectorAll('elements table row cell')];
    assert.equal(cells.length, 3);
    for (const cell of cells.slice(1)) {
        const paragraphs = cell.querySelectorAll('paragraph');
        assert.equal(paragraphs.length, 1);
        assert.equal(paragraphs[0].getAttribute('TabSet'), null);
        assert.deepEqual([...paragraphs[0].children].map(el => el.getAttribute('length')), ['0']);
    }
});

test('a footnote, a picture and a nested table in a joined cell give no findings either', async () => {
    const { parsed, content } = await convert(await buildDocx(mergeColumn([
        ['restart', p(r('Birleşik'))],
        ['continue', p(r('Dipnotlu') + footnoteRef(1)) + p(r('Resimli') + picture('rId_yok')) +
            p(picture()) + tbl([[lossy('İç tablo')]])]
    ]), { footnotes: { 1: p(r('Yazılmayan dipnot')) }, media: { 'image1.png': IMAGES.gif } }));

    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.equal(content.trim(), 'Birleşik');
});

test('a continuation cell with no merge above it is converted, so its findings count', async () => {
    const { parsed, content } = await convert(await buildDocx(mergeColumn([['continue', lossy('Sahipsiz devam')]])));

    assert.ok(content.includes('Sahipsiz devam'), JSON.stringify(content));
    assert.deepEqual(findingSummaries(parsed.findings), lossyFindings('Sahipsiz devam'));
});

test('a nested table\'s joined cell is written empty too, so its findings don\'t count', async () => {
    const nested = mergeColumn([['restart', p(r('İç başlık'))], ['continue', lossy('Yazılmayan iç devam')]]);
    const { parsed, content } = await convert(await buildDocx(tbl([[nested + p(''), p(r('Yan hücre'))]])));

    assert.ok(!content.includes('Yazılmayan iç devam'), JSON.stringify(content));
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('a joined nested cell doesn\'t end the outer joined cell: what follows it in that cell gives no findings', async () => {
    const nested = mergeColumn([['restart', p(r('İç'))], ['continue', p(r('İç devam'))]]);
    const { parsed, content } = await convert(await buildDocx(mergeColumn([
        ['restart', p(r('Birleşik'))],
        ['continue', nested + lossy('Sonra yazılmayan')]
    ])));

    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.equal(content.trim(), 'Birleşik');
});

test('a table nested past the limit is flattened with its joined cells, so their findings count', async () => {
    const nested = mergeColumn([['restart', p(r('İç başlık'))], ['continue', lossy('Düzleştirilen devam')]]);
    const { parsed, content } = await convert(await buildDocx(tbl([[inTwoTables(nested), p(r('Yan hücre'))]])));

    assert.ok(content.includes('Düzleştirilen devam'), JSON.stringify(content));
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'nestedTables', count: 1, examples: ['İç başlık'] },
        ...lossyFindings('Düzleştirilen devam')
    ]);
});

const pageBreak = '<w:r><w:br w:type="page"/></w:r>';
// A paragraph of these runs with an exact line height, a bar stop and a left stop: converted, it
// gives one lineSpacing and one tabStops finding and is written with TabSet="72.0:0:0"
const lossyRuns = runs => `<w:p><w:pPr>${spacing('exact') + tabsOf(stop('bar', 720), stop('left', 1440))}</w:pPr>${runs}</w:p>`;
const writtenLossy = doc => doc.querySelectorAll('elements paragraph[TabSet="72.0:0:0"]').length;
const topLevel = doc => [...doc.querySelector('elements').children].map(el => el.localName);

test('a body paragraph holding only a page break is written as the break alone, and gives no findings', async () => {
    const { parsed, doc } = await convert(await buildDocx(
        p(r('Önce')) + lossyRuns(pageBreak) + p(r('Arada')) + lossyRuns(pageBreak + pageBreak) + p(r('Sonra'))));

    assert.deepEqual(topLevel(doc), ['paragraph', 'page-break', 'paragraph', 'page-break', 'page-break', 'paragraph']);
    assert.equal(writtenLossy(doc), 0);
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('a picture that is skipped does not make a page-break paragraph a written one', async () => {
    const { parsed, doc } = await convert(await buildDocx(p(r('Önce')) + lossyRuns(picture('rId_yok') + pageBreak)));

    assert.deepEqual(topLevel(doc), ['paragraph', 'page-break']);
    // The picture is still reported as skipped, with the nearest text
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'missingImage', count: 1, examples: ['Önce'] }]);
});

test('a paragraph with text before or after its page break is written, and counted once', async () => {
    const { parsed, doc } = await convert(await buildDocx(
        lossyRuns(r('Kesmeden önce') + pageBreak) + lossyRuns(pageBreak + r('Kesmeden sonra')) +
        lossyRuns(r('İki') + pageBreak + r('yanında'))));

    assert.equal(writtenLossy(doc), 4, 'the paragraph split in two is written as two');
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'lineSpacing', count: 3, examples: ['Kesmeden önce', 'Kesmeden sonra', 'İkiyanında'] },
        { code: 'tabStops', count: 3, examples: ['Kesmeden önce', 'Kesmeden sonra', 'İkiyanında'] }
    ]);
});

test('in a table cell a page break is a line break, so a paragraph holding only one is written and counted', async () => {
    const { parsed, doc } = await convert(await buildDocx(tbl([[lossyRuns(pageBreak), p(r('Yan hücre'))]])));

    assert.equal(doc.querySelector('elements page-break'), null);
    assert.equal(writtenLossy(doc), 1);
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'lineSpacing', count: 1, examples: [] },
        { code: 'tabStops', count: 1, examples: [] }
    ]);
});

test('a text box paragraph in the body follows the body rule, one in a cell the cell rule', async () => {
    const textBox = inner => '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

    const inBody = await convert(await buildDocx(p(r('Çapa') + textBox(lossyRuns(pageBreak)))));
    assert.equal(writtenLossy(inBody.doc), 0);
    assert.deepEqual(findingSummaries(inBody.parsed.findings), []);

    const inCell = await convert(await buildDocx(tbl([[p(r('Çapa') + textBox(lossyRuns(pageBreak)))]])));
    assert.equal(writtenLossy(inCell.doc), 1);
    assert.deepEqual(findingSummaries(inCell.parsed.findings).map(finding => [finding.code, finding.count]),
        [['lineSpacing', 1], ['tabStops', 1]]);
});

// Every kind of item a run can hold, and whether it puts something in the converted text
const RUN_ITEMS = {
    text: [r('Metin'), true],
    emptyText: ['<w:r><w:t></w:t></w:r>', false],
    tab: ['<w:r><w:tab/></w:r>', true],
    lineBreak: ['<w:r><w:br/></w:r>', true],
    carriageReturn: ['<w:r><w:cr/></w:r>', true],
    picture: [picture(), true],
    missingPicture: [picture('rId_yok'), false],
    footnote: [footnoteRef(1), true],
    endnote: [endnoteRef(1), true]
};

test('lineSpacing and tabStops are recorded for exactly the paragraphs the generator writes', async () => {
    const cases = [[], ['pageBreak'], ['pageBreak', 'pageBreak']];
    for (const name of Object.keys(RUN_ITEMS)) {
        cases.push([name], [name, 'pageBreak'], ['pageBreak', name], ['pageBreak', name, 'pageBreak']);
    }

    for (const context of ['body', 'cell']) {
        for (const names of cases) {
            const label = `${context}: [${names.join(', ')}]`;
            const paragraph = lossyRuns(names.map(name => (name === 'pageBreak' ? pageBreak : RUN_ITEMS[name][0])).join(''));
            const { parsed, doc } = await convert(await buildDocx(context === 'cell' ? tbl([[paragraph]]) : paragraph, {
                footnotes: { 1: p(r('Dipnot')) }, endnotes: ENDNOTES, media: { 'image1.png': IMAGES.png }
            }));

            // What the generator did with the paragraph, and what it is expected to do
            const written = writtenLossy(doc) > 0;
            const hasContent = names.some(name => name !== 'pageBreak' && RUN_ITEMS[name][1]);
            assert.equal(written, context === 'cell' || hasContent || !names.includes('pageBreak'), `${label} written`);

            // The parser's findings follow it
            const counts = Object.fromEntries(findingSummaries(parsed.findings).map(finding => [finding.code, finding.count]));
            assert.equal(counts.lineSpacing || 0, written ? 1 : 0, `${label} lineSpacing`);
            assert.equal(counts.tabStops || 0, written ? 1 : 0, `${label} tabStops`);
        }
    }
});

test('paragraphs and footnotes that are never converted give no finding', async () => {
    // A deleted footnote reference, and tab stops and line heights only in a style no paragraph uses
    const findings = await findingsOf(
        p(r('Metin') + `<w:del w:id="1" w:author="Y">${footnoteRef(1)}</w:del>`),
        { footnotes: { 1: p(r('Silinen')) }, styles: paraStyle('Unused', spacing('exact') + tabsOf(stop('bar', 720))) });
    assert.deepEqual(findings, []);
});

test('the new findings come in FINDING_CODES order after the existing ones', async () => {
    const body = sectionBreak() + withPPr(tabsOf(stop('bar', 720)), 'Çubuk') + withPPr(spacing('exact'), 'Tam') +
        p(r('Son not') + endnoteRef(1)) + p(r('Dipnot') + footnoteRef(1)) + p(r('Resim') + picture('rId_yok'));
    const findings = await findingsOf(body, {
        footnotes: { 1: p(r('Dipnot metni')) },
        endnotes: ENDNOTES,
        settings: endnotePr(pos('sectEnd'))
    });
    assert.deepEqual(findings.map(finding => finding.code),
        ['missingImage', 'notesMoved', 'endnotesMoved', 'lineSpacing', 'tabStops']);
});

test('the UDF of a document with every new finding is the same with or without its findings', async () => {
    const { DocxParser, UdfGenerator } = loadConverter();
    const docx = await buildDocx(sectionBreak() + withPPr(tabsOf(stop('bar', 720), stop('left', 1440, 'heavy')), 'Çubuk') +
        withPPr(spacing('exact'), 'Tam') + p(r('Son not') + endnoteRef(1)) + p(r('Dipnot') + footnoteRef(1)),
    { footnotes: { 1: p(r('Dipnot metni')) }, endnotes: ENDNOTES, settings: endnotePr(pos('sectEnd')) });
    const parsed = await new DocxParser().parse(docx);
    assert.equal(parsed.findings.length, 4);
    const udfXml = async document => (await JSZip.loadAsync(
        await (await new UdfGenerator().generate(document)).arrayBuffer())).file('content.xml').async('string');
    assert.equal(await udfXml(parsed), await udfXml({ ...parsed, findings: [] }));
});

// Locations: where each occurrence of a finding is

// Each finding of a parse as plain data in this realm, its occurrences included
async function locatedFindingsOf(bodyXml, options) {
    const { DocxParser } = loadConverter();
    return plain((await new DocxParser().parse(await buildDocx(bodyXml, options))).findings);
}
const occurrencesOf = (findings, code) => (findings.find(finding => finding.code === code) || { occurrences: [] }).occurrences;
const locationsOf = (findings, code) => occurrencesOf(findings, code).map(occurrence => occurrence.location);

const hiddenRun = text => `<w:r><w:rPr><w:vanish/></w:rPr><w:t>${text}</w:t></w:r>`;
const bold = text => `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const exact = text => withPPr(spacing('exact'), text);
const vmlTextBox = inner => `<w:r><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
const HEADING_STYLES = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Heading1"/>' +
    '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>';
const styled = (styleId, text) => withPPr(`<w:pStyle w:val="${styleId}"/>`, text);

test('every finding lists its occurrences, at most 5, each with an example and a location, and keeps code, count and examples', async () => {
    for (const petition of petitions) {
        const { DocxParser } = loadConverter();
        const findings = plain((await new DocxParser().parse(await petition.build())).findings);
        for (const finding of findings) {
            assert.deepEqual(Object.keys(finding), ['code', 'count', 'examples', 'occurrences'], petition.name);
            assert.ok(finding.occurrences.length >= 1 && finding.occurrences.length <= 5, `${petition.name} ${finding.code}`);
            for (const occurrence of finding.occurrences) {
                assert.deepEqual(Object.keys(occurrence), ['example', 'location']);
                assert.equal(typeof occurrence.example, 'string');
                assert.ok(['paragraph', 'cell', 'header', 'footer', 'document'].includes(occurrence.location.kind));
            }
        }
    }
});

test('body paragraphs are numbered from 1, empty ones included, and an extra paragraph before a finding moves it by one', async () => {
    const body = p('') + p(r('Birinci')) + p('') + exact('Tam satır');
    assert.deepEqual(occurrencesOf(await locatedFindingsOf(body), 'lineSpacing'),
        [{ example: 'Tam satır', location: { kind: 'paragraph', paragraph: 4 } }]);
    assert.deepEqual(locationsOf(await locatedFindingsOf(p(r('Eklenen')) + body), 'lineSpacing'),
        [{ kind: 'paragraph', paragraph: 5 }]);
    // Tables don't count as paragraphs, and content controls are looked through
    const withTable = tbl([[p(r('Hücre')), p('')]]) + `<w:sdt><w:sdtContent>${p(r('Denetimde'))}</w:sdtContent></w:sdt>` + exact('Sonra');
    assert.deepEqual(locationsOf(await locatedFindingsOf(withTable), 'lineSpacing'), [{ kind: 'paragraph', paragraph: 2 }]);
});

test('a finding in a cell names its table, row and cell; nested tables report the outermost cell', async () => {
    const third = [p(r('Görünen') + hiddenRun('Gizli hücre')) +
        inTwoTables(tbl([[p(r('İç') + hiddenRun('İç gizli')), p('')]])), p('')];
    const body = p(r('Giriş')) + tbl([[p(r('Bir')), p(r('İki'))]]) + p('') +
        tbl([[p(r('A')), p(r('B'))], [p(r('C')), exact('D')], third]);
    const findings = await locatedFindingsOf(body);
    const cell = (row, cellNumber) => ({ kind: 'cell', table: 2, row, cell: cellNumber });
    assert.deepEqual(occurrencesOf(findings, 'hiddenText'), [
        { example: 'Gizli hücre', location: cell(3, 1) },
        { example: 'İç gizli', location: cell(3, 1) }
    ]);
    assert.deepEqual(occurrencesOf(findings, 'nestedTables'), [{ example: 'İç', location: cell(3, 1) }]);
    assert.deepEqual(occurrencesOf(findings, 'lineSpacing'), [{ example: 'D', location: cell(2, 2) }]);
});

test('a text box\'s findings carry its anchoring paragraph, and its paragraphs don\'t shift the numbering', async () => {
    const body = p(r('Birinci')) +
        p(r('Çapa') + vmlTextBox(p(r('Kutu') + hiddenRun('Kutuda gizli')) + exact('Kutuda tam') + p(r('Üçüncü kutu paragrafı')))) +
        exact('Sonraki');
    const findings = await locatedFindingsOf(body);
    assert.deepEqual(locationsOf(findings, 'hiddenText'), [{ kind: 'paragraph', paragraph: 2 }]);
    assert.deepEqual(occurrencesOf(findings, 'lineSpacing'), [
        { example: 'Kutuda tam', location: { kind: 'paragraph', paragraph: 2 } },
        { example: 'Sonraki', location: { kind: 'paragraph', paragraph: 3 } }
    ]);
});

test('the heading is the last outline-level paragraph before the place, from a Heading style or a direct w:outlineLvl', async () => {
    const body = exact('Başlıktan önce') +
        styled('Heading1', 'Olaylar') + exact('Olaylarda') +
        withPPr('<w:outlineLvl w:val="2"/>', 'Doğrudan düzey') + exact('Düzeyde') +
        styled('Heading2', 'Alt başlık') + tbl([[exact('Tabloda'), p('')]]);
    const findings = await locatedFindingsOf(body, { styles: HEADING_STYLES });
    assert.deepEqual(locationsOf(findings, 'lineSpacing'), [
        { kind: 'paragraph', paragraph: 1 },
        { kind: 'paragraph', paragraph: 3, heading: 'Olaylar' },
        { kind: 'paragraph', paragraph: 5, heading: 'Doğrudan düzey' },
        { kind: 'cell', table: 1, row: 1, cell: 1, heading: 'Alt başlık' }
    ]);
});

test('outline level 9 (body text) is no heading, even when the paragraph sets it over a Heading style', async () => {
    const body = styled('Heading1', 'Olaylar') + styled('BodyText', 'Gövde metni düzeyi') + exact('Gövdeden sonra') +
        withPPr('<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>', 'Düzeyi kaldırılmış') + exact('Kaldırılandan sonra');
    assert.deepEqual(locationsOf(await locatedFindingsOf(body, { styles: HEADING_STYLES }), 'lineSpacing'), [
        { kind: 'paragraph', paragraph: 3, heading: 'Olaylar' },
        { kind: 'paragraph', paragraph: 5, heading: 'Olaylar' }
    ]);
});

test('a heading\'s own finding names the heading before it, and a long one is shortened', async () => {
    const long = 'Uzun bir başlık metni kırk karakteri rahatlıkla geçer ve kısaltılır';
    const body = styled('Heading1', 'Birinci başlık') + withPPr('<w:pStyle w:val="Heading1"/>' + spacing('exact'), 'İkinci başlık') +
        styled('Heading1', long) + exact('Altında');
    const findings = await locatedFindingsOf(body, { styles: HEADING_STYLES });
    assert.deepEqual(locationsOf(findings, 'lineSpacing'), [
        { kind: 'paragraph', paragraph: 2, heading: 'Birinci başlık' },
        { kind: 'paragraph', paragraph: 4, heading: 'Uzun bir başlık metni kırk karakteri rah…' }
    ]);
});

test('without outline levels, a short all-bold or all-upper-case body paragraph is the heading', async () => {
    const body = exact('Önce') +
        p(r('AÇIKLAMALAR :')) + exact('Açıklamada') +
        p(bold('Deliller') + r(' ') + bold('ve ekler')) + exact('Delillerde') +
        p(r('SONUÇ VE İSTEM') + '<w:r><w:t xml:space="preserve">  </w:t></w:r>') + exact('Sonuçta') +
        p(r('Karışık ') + bold('kalın')) + p(r('X'.repeat(81))) + p(r('2024/123')) + exact('Başlık değil') +
        tbl([[p(bold('Hücrede kalın')), p('')]]) + exact('Tablodan sonra');
    const findings = await locatedFindingsOf(body);
    assert.deepEqual(locationsOf(findings, 'lineSpacing').map(location => location.heading), [
        undefined, 'AÇIKLAMALAR :', 'Deliller ve ekler', 'SONUÇ VE İSTEM', 'SONUÇ VE İSTEM'
    ]);
    // A bold paragraph in a table is no heading
    const tail = await locatedFindingsOf(p(r('SONUÇ')) + tbl([[p(bold('Hücrede kalın')), p('')]]) + exact('Tablodan sonra'));
    assert.deepEqual(locationsOf(tail, 'lineSpacing'), [{ kind: 'paragraph', paragraph: 2, heading: 'SONUÇ' }]);
});

test('an outline-level heading wins over a later bold or upper-case one', async () => {
    const body = styled('Heading1', 'Olaylar') + p(r('AÇIKLAMALAR')) + exact('Altında');
    assert.deepEqual(locationsOf(await locatedFindingsOf(body, { styles: HEADING_STYLES }), 'lineSpacing'),
        [{ kind: 'paragraph', paragraph: 3, heading: 'Olaylar' }]);
});

test('note text findings and moved footnotes are located at the reference, marked with the note kind', async () => {
    const sym = '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>';
    const body = p(r('BAŞLIK')) + p(r('Gövde') + footnoteRef(1)) + tbl([[p(r('Hücre') + endnoteRef(1)), p('')]]);
    const findings = await locatedFindingsOf(body, {
        footnotes: { 1: p(r('Dipnot') + sym + hiddenRun('Dipnotta gizli')) },
        endnotes: { 1: p(r('Son not') + sym) }
    });
    const inFootnote = { kind: 'paragraph', paragraph: 2, heading: 'BAŞLIK', note: { kind: 'footnote' } };
    const inEndnote = { kind: 'cell', table: 1, row: 1, cell: 1, heading: 'BAŞLIK', note: { kind: 'endnote' } };
    assert.deepEqual(locationsOf(findings, 'notesMoved'), [inFootnote]);
    assert.deepEqual(locationsOf(findings, 'hiddenText'), [inFootnote]);
    assert.deepEqual(locationsOf(findings, 'symbols'), [inFootnote, inEndnote]);
});

test('header and footer findings are located in them; headersFooters, comments and endnotesMoved are document-wide', async () => {
    const sym = '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>';
    const findings = await locatedFindingsOf(sectionBreak() + p(r('Son not') + endnoteRef(1)), {
        sectPr: headerRefs + '<w:headerReference w:type="first" r:id="rIdHeader2"/>',
        endnotes: ENDNOTES,
        settings: endnotePr(pos('sectEnd')),
        parts: [
            header('header1.xml', 'rIdHeader1', p(r('Üst') + hiddenRun('Üstte gizli'))),
            header('footer1.xml', 'rIdFooter1', p(r('Alt') + sym), 'footer'),
            header('header2.xml', 'rIdHeader2', p(r('İlk sayfa'))),
            comments(['Yorum'])
        ]
    });
    assert.deepEqual(locationsOf(findings, 'hiddenText'), [{ kind: 'header' }]);
    assert.deepEqual(locationsOf(findings, 'symbols'), [{ kind: 'footer' }]);
    assert.deepEqual(occurrencesOf(findings, 'headersFooters'), [{ example: '', location: { kind: 'document' } }]);
    assert.deepEqual(occurrencesOf(findings, 'comments'), [{ example: '', location: { kind: 'document' } }]);
    assert.deepEqual(occurrencesOf(findings, 'endnotesMoved'), [{ example: 'Birinci son not', location: { kind: 'document' } }]);
});

test('occurrences are capped at 5 in document order, and an example repeated at one place is listed once', async () => {
    const sym = '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>';
    const body = p(sym + sym) + Array.from({ length: 7 }, () => exact('Aynı')).join('') + p(sym);
    const findings = await locatedFindingsOf(body);
    const lineSpacing = findings.find(finding => finding.code === 'lineSpacing');
    assert.equal(lineSpacing.count, 7);
    assert.deepEqual(lineSpacing.examples, ['Aynı']);
    assert.deepEqual(lineSpacing.occurrences.map(({ example, location }) => [example, location.paragraph]),
        [['Aynı', 2], ['Aynı', 3], ['Aynı', 4], ['Aynı', 5], ['Aynı', 6]]);
    const symbols = findings.find(finding => finding.code === 'symbols');
    assert.equal(symbols.count, 3);
    assert.deepEqual(symbols.occurrences.map(({ location }) => location.paragraph), [1, 9]);
});

test('a drawing keeps its nearest-paragraph example and takes the place of the paragraph or cell holding it', async () => {
    const body = p(r('Önce')) + p(picture('rId_yok')) + tbl([[p(r('Hücre')), p(picture('rId_yok2'))]]);
    assert.deepEqual(occurrencesOf(await locatedFindingsOf(body), 'missingImage'), [
        { example: 'Önce', location: { kind: 'paragraph', paragraph: 2 } },
        { example: 'Hücre', location: { kind: 'cell', table: 1, row: 1, cell: 2 } }
    ]);
});

test('a cell joined to the merge above it still records no occurrence', async () => {
    const findings = await locatedFindingsOf(mergeColumn([['restart', exact('Birleşik')], ['continue', exact('Yazılmayan')]]));
    assert.deepEqual(occurrencesOf(findings, 'lineSpacing'),
        [{ example: 'Birleşik', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }]);
});
