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

test('a page number field, a table or a picture in the converted header or footer gives the finding', async () => {
    const pageField = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const cases = {
        'PAGE field': ['footer', p(r('Sayfa ') + pageField)],
        'NUMPAGES simple field': ['footer', p(r('Toplam ') + '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple>')],
        table: ['header', tbl([[p(r('Hücre')), p('')]]) + p(r('Üst bilgi'))],
        picture: ['header', p(r('Üst bilgi') + picture('rIdLogo'))]
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
    const docx = await buildDocx(tbl([[p(r('Dış')) + tbl([[p(r('İç'))]])]]) + p(picture()),
        { media: { 'image1.png': IMAGES.gif } });
    const parsed = await new DocxParser().parse(docx);
    assert.equal(parsed.findings.length, 2);
    const udfXml = async document => (await JSZip.loadAsync(
        await (await new UdfGenerator().generate(document)).arrayBuffer())).file('content.xml').async('string');
    assert.equal(await udfXml(parsed), await udfXml({ ...parsed, findings: [] }));
});

// None of the petition fixtures has headers, comments, nested tables, pictures, endnotes, exact or
// at-least line heights, or bar stops and underscore or heavy leaders. Only the dava dilekçesi has footnotes.
const EXPECTED_FIXTURE_FINDINGS = {
    'dava dilekçesi': [{
        code: 'notesMoved',
        count: 2,
        examples: ['Ek-1: Hizmet sözleşmesinin örneği.', 'Ek-2: Örnek Noterliğinin 00.00.2026 tari…']
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
    assert.deepEqual(plain(parsed.findings), lossyFindings('Birleşik'));
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

    assert.deepEqual(plain(parsed.findings), []);
    assert.equal(content.trim(), 'Birleşik');
});

test('a continuation cell with no merge above it is converted, so its findings count', async () => {
    const { parsed, content } = await convert(await buildDocx(mergeColumn([['continue', lossy('Sahipsiz devam')]])));

    assert.ok(content.includes('Sahipsiz devam'), JSON.stringify(content));
    assert.deepEqual(plain(parsed.findings), lossyFindings('Sahipsiz devam'));
});

test('a nested table is flattened with its joined cells, so their findings count', async () => {
    const nested = mergeColumn([['restart', p(r('İç başlık'))], ['continue', lossy('Düzleştirilen devam')]]);
    const { parsed, content } = await convert(await buildDocx(tbl([[nested, p(r('Yan hücre'))]])));

    assert.ok(content.includes('Düzleştirilen devam'), JSON.stringify(content));
    assert.deepEqual(plain(parsed.findings), [
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
    assert.deepEqual(plain(parsed.findings), []);
});

test('a picture that is skipped does not make a page-break paragraph a written one', async () => {
    const { parsed, doc } = await convert(await buildDocx(p(r('Önce')) + lossyRuns(picture('rId_yok') + pageBreak)));

    assert.deepEqual(topLevel(doc), ['paragraph', 'page-break']);
    // The picture is still reported as skipped, with the nearest text
    assert.deepEqual(plain(parsed.findings), [{ code: 'missingImage', count: 1, examples: ['Önce'] }]);
});

test('a paragraph with text before or after its page break is written, and counted once', async () => {
    const { parsed, doc } = await convert(await buildDocx(
        lossyRuns(r('Kesmeden önce') + pageBreak) + lossyRuns(pageBreak + r('Kesmeden sonra')) +
        lossyRuns(r('İki') + pageBreak + r('yanında'))));

    assert.equal(writtenLossy(doc), 4, 'the paragraph split in two is written as two');
    assert.deepEqual(plain(parsed.findings), [
        { code: 'lineSpacing', count: 3, examples: ['Kesmeden önce', 'Kesmeden sonra', 'İkiyanında'] },
        { code: 'tabStops', count: 3, examples: ['Kesmeden önce', 'Kesmeden sonra', 'İkiyanında'] }
    ]);
});

test('in a table cell a page break is a line break, so a paragraph holding only one is written and counted', async () => {
    const { parsed, doc } = await convert(await buildDocx(tbl([[lossyRuns(pageBreak), p(r('Yan hücre'))]])));

    assert.equal(doc.querySelector('elements page-break'), null);
    assert.equal(writtenLossy(doc), 1);
    assert.deepEqual(plain(parsed.findings), [
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
    assert.deepEqual(plain(inBody.parsed.findings), []);

    const inCell = await convert(await buildDocx(tbl([[p(r('Çapa') + textBox(lossyRuns(pageBreak)))]])));
    assert.equal(writtenLossy(inCell.doc), 1);
    assert.deepEqual(plain(inCell.parsed.findings).map(finding => [finding.code, finding.count]),
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
            const counts = Object.fromEntries(plain(parsed.findings).map(finding => [finding.code, finding.count]));
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
