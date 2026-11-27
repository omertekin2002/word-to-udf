const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { buildDocx, convert, convertBody, textOf, findingSummaries } = require('./helpers');

// Links (w:hyperlink and HYPERLINK fields). Docs.md defines no link element, so only a link's text
// is written; a link whose text doesn't show its address is reported as a hyperlinks conversion note.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
/** A w:hyperlink around these runs, through relationship id (or to a bookmark, with anchor) */
const hyperlink = (id, runs, anchor = '') =>
    `<w:hyperlink${id ? ` r:id="${id}"` : ''}${anchor ? ` w:anchor="${anchor}"` : ''} w:history="1">${runs}</w:hyperlink>`;
const fldChar = type => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = text => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
/** A complex field: begin, its instruction (one w:instrText run per piece), separate, the result runs, end */
const field = (instructions, result) =>
    fldChar('begin') + [].concat(instructions).map(instr).join('') + fldChar('separate') + result + fldChar('end');
const fldSimple = (instruction, result) => `<w:fldSimple w:instr="${instruction}">${result}</w:fldSimple>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const textBox = inner => '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const footnoteRef = id => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const HEADING_STYLE = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';

const KARAR = 'https://example.org/karar';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;

/** Text of each body paragraph, from the CDATA pool */
const bodyParagraphs = (doc, content) => [...doc.querySelectorAll('elements > paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));

test('the issue\'s reproduction: an external w:hyperlink and a HYPERLINK field are reported with their addresses', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Karar için bkz. ') + hyperlink('rIdLink1', r('Yargıtay kararı', '<w:rStyle w:val="Hyperlink"/>'))) +
        p(r('Field: ') + field(['HYPERLINK "', `${KARAR}" `], r('buradan'))),
        { links: { 'document.xml': { rIdLink1: KARAR } } });

    // Only the links' text is written
    assert.deepEqual(bodyParagraphs(doc, content), ['Karar için bkz. Yargıtay kararı', 'Field: buradan']);
    assert.equal(content.includes('example.org'), false);
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'hyperlinks', count: 2, examples: ['Yargıtay kararı → example.org/karar', 'buradan → example.org/karar'] }
    ]);
    assert.deepEqual(occurrences(parsed, 'hyperlinks'), [
        { example: 'Yargıtay kararı → example.org/karar', location: { kind: 'paragraph', paragraph: 1 } },
        { example: 'buradan → example.org/karar', location: { kind: 'paragraph', paragraph: 2 } }
    ]);
});

test('a w:fldSimple HYPERLINK, an e-mail link and a field with switches are reported the same way', async () => {
    const { content, parsed } = await convertBody(
        p(r('e-Devlet: ') + fldSimple(`HYPERLINK &quot;https://www.turkiye.gov.tr/&quot;`, r('buradan erişilebilir'))) +
        p(r('Yazışma: ') + hyperlink('rIdMail', r('büroya yazın'))) +
        p(field(' HYPERLINK "https://example.org/kanun" \\l "madde5" \\o "İpucu" ', r('kanun metni'))) +
        // Lower-case field name and an unquoted address
        p(field('hyperlink https://example.org/yonetmelik', r('yönetmelik'))),
        { links: { 'document.xml': { rIdMail: 'mailto:buro@example.org' } } });

    assert.equal(content.includes('buradan erişilebilir'), true);
    assert.deepEqual(occurrences(parsed, 'hyperlinks').map(o => [o.example, o.location.paragraph]), [
        // Shortened like any example
        ['buradan erişilebilir → www.turkiye.gov.t…', 1],
        ['büroya yazın → buro@example.org', 2],
        ['kanun metni → example.org/kanun', 3],
        ['yönetmelik → example.org/yonetmelik', 4]
    ]);
});

test('a link whose text shows its address, an internal link, an empty or hidden or deleted link, or a missing relationship is not reported', async () => {
    const links = {
        rIdSame: 'https://example.org/karar',
        rIdWww: 'https://www.example.org/',
        rIdMail: 'mailto:Buro@Example.org',
        rIdHidden: KARAR,
        rIdDeleted: KARAR,
        rIdEmpty: KARAR,
        rIdSpace: KARAR
    };
    const { content, parsed } = await convertBody(
        // The text is the address, with or without its scheme, www. or trailing slash, in any case
        p(hyperlink('rIdSame', r('https://example.org/karar'))) +
        p(hyperlink('rIdSame', r('EXAMPLE.ORG/KARAR/'))) +
        p(hyperlink('rIdWww', r('www.example.org'))) +
        p(hyperlink('rIdWww', r('Adres: example.org, bkz.'))) +
        p(hyperlink('rIdMail', r('buro@example.org'))) +
        p(field('HYPERLINK "http://www.example.org/karar/"', r('https://example.org/karar'))) +
        // Internal links
        p(hyperlink('', r('Ek-1'), 'Ek1')) +
        p(field('HYPERLINK \\l "Ek2"', r('Ek-2'))) +
        p(fldSimple('HYPERLINK \\l &quot;Ek3&quot;', r('Ek-3'))) +
        // Other field types
        p(field('REF Ek4 \\h', r('Ek-4')) + field('PAGEREF Ek4', r('3'))) +
        // Nothing shown: all hidden, all deleted, no text, only whitespace
        p(r('Gizli: ') + hyperlink('rIdHidden', r('gizli', '<w:vanish/>'))) +
        p(r('Silinen: ') + hyperlink('rIdDeleted',
            '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>silindi</w:delText></w:r></w:del>')) +
        p(r('Boş: ') + hyperlink('rIdEmpty', '') + hyperlink('rIdSpace', r(' '))) +
        // A relationship that doesn't exist
        p(r('Kayıp: ') + hyperlink('rIdMissing', r('kayıp bağlantı'))) +
        // A field with no result
        p(fldChar('begin') + instr(`HYPERLINK "${KARAR}"`) + fldChar('end')),
        { links: { 'document.xml': links } });

    assert.equal(content.includes('kayıp bağlantı'), true);
    assert.equal(finding(parsed, 'hyperlinks'), undefined);
});

test('a link to a document part (no TargetMode) is not an external link', async () => {
    // A w:hyperlink whose r:id names an internal relationship, such as the styles part
    const { parsed } = await convertBody(p(hyperlink('rId_styles', r('Stiller'))), { styles: '' });
    assert.equal(finding(parsed, 'hyperlinks'), undefined);
});

test('two links in one paragraph count 2, and a link split over several runs is one link', async () => {
    const { parsed } = await convertBody(
        p(hyperlink('rIdA', r('Yargıtay ') + r('kararı', '<w:b/>')) + r(' ve ') +
            hyperlink('rIdB', r('Danıştay kararı'))),
        { links: { 'document.xml': { rIdA: 'https://example.org/a', rIdB: 'https://example.org/b' } } });

    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'hyperlinks', count: 2, examples: ['Yargıtay kararı → example.org/a', 'Danıştay kararı → example.org/b'] }
    ]);
    assert.deepEqual(occurrences(parsed, 'hyperlinks').map(o => o.location), [
        { kind: 'paragraph', paragraph: 1 }, { kind: 'paragraph', paragraph: 1 }
    ]);
});

test('a field nested in a link field\'s result is read, and the outer link keeps the nested result\'s text', async () => {
    // HYPERLINK "…/karar" whose result holds a REF field and a HYPERLINK field of its own
    const { parsed } = await convertBody(p(
        fldChar('begin') + instr(`HYPERLINK "${KARAR}"`) + fldChar('separate') +
        r('Karar ') + field('REF Esas', r('2023/1')) + r(' ') + field('HYPERLINK "https://example.org/ek"', r('ek')) +
        fldChar('end')));

    assert.deepEqual(occurrences(parsed, 'hyperlinks').map(o => o.example), [
        'Karar 2023/1 ek → example.org/karar',
        'ek → example.org/ek'
    ]);
});

test('a hidden part of a link\'s text is left out of its example; a hidden field code still makes it a link', async () => {
    const { parsed } = await convertBody(p(
        hyperlink('rIdLink1', r('Yargıtay') + r(' gizli', '<w:vanish/>') + r(' kararı')) +
        fldChar('begin') + `<w:r><w:rPr><w:vanish/></w:rPr><w:instrText>HYPERLINK "https://example.org/ek"</w:instrText></w:r>` +
        fldChar('separate') + r(' ek') + fldChar('end')),
        { links: { 'document.xml': { rIdLink1: KARAR } } });

    assert.deepEqual(occurrences(parsed, 'hyperlinks').map(o => o.example),
        ['Yargıtay kararı → example.org/karar', 'ek → example.org/ek']);
});

test('a long example is shortened like the other notes\' examples, with one address', async () => {
    const { parsed } = await convertBody(
        p(hyperlink('rIdLink1', r('Yargıtay 3. Hukuk Dairesinin 2023/1234 Esas sayılı kararı'))),
        { links: { 'document.xml': { rIdLink1: KARAR } } });
    const [example] = finding(parsed, 'hyperlinks').examples;
    assert.equal(example, 'Yargıtay 3. Hukuk Dairesinin 2023/1234 E…');
});

test('links in a table cell, a text box, the header and footer, and a footnote are reported with their locations', async () => {
    const parts = [
        { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Büro ') + hyperlink('rIdSite', r('web sitesi'))) },
        { name: 'footer1.xml', root: 'ftr', type: 'footer', id: 'rIdFooter1', inner: p(r('İletişim: ') + hyperlink('rIdMail', r('e-posta'))) }
    ];
    const sectPr = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';
    const links = {
        // The header's and footer's ids resolve through their own parts, not the document's
        'document.xml': { rIdCell: 'https://example.org/delil', rIdBox: 'https://example.org/kutu', rIdSite: 'https://example.org/yanlis' },
        'header1.xml': { rIdSite: 'https://example.org/buro' },
        'footer1.xml': { rIdMail: 'mailto:buro@example.org' },
        'footnotes.xml': { rIdNote: 'https://example.org/kaynak' }
    };
    const { parsed } = await convertBody(
        p(r('DELİLLER'), '<w:pStyle w:val="Heading1"/>') +
        tbl([p(r('A')), p(r('Bkz. ') + hyperlink('rIdCell', r('delil')))]) +
        p(r('Kutulu') + textBox(p(hyperlink('rIdBox', r('kutudaki bağlantı'))))) +
        p(r('Dipnot') + footnoteRef(1)),
        { styles: HEADING_STYLE, parts, sectPr, links, footnotes: { 1: p(r('Kaynak: ') + hyperlink('rIdNote', r('karar metni'))) } });

    const found = occurrences(parsed, 'hyperlinks');
    const at = example => found.filter(o => o.example === example).map(o => o.location);
    assert.deepEqual(at('delil → example.org/delil'), [{ kind: 'cell', table: 1, row: 1, cell: 2, heading: 'DELİLLER' }]);
    // The table isn't a body paragraph; the text box takes the place of the paragraph anchoring it
    assert.deepEqual(at('kutudaki bağlantı → example.org/kutu'), [{ kind: 'paragraph', paragraph: 2, heading: 'DELİLLER' }]);
    assert.deepEqual(at('web sitesi → example.org/buro'), [{ kind: 'header' }]);
    assert.deepEqual(at('e-posta → buro@example.org'), [{ kind: 'footer' }]);
    assert.deepEqual(at('karar metni → example.org/kaynak'),
        [{ kind: 'paragraph', paragraph: 3, heading: 'DELİLLER', note: { kind: 'footnote' } }]);
    assert.equal(finding(parsed, 'hyperlinks').count, 5);
});

test('a link in a footnote or endnote is reported once per reference to the note', async () => {
    const endnoteRef = id => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;
    const { parsed } = await convertBody(
        p(r('Dipnot') + footnoteRef(1)) + p(r('Yine') + footnoteRef(1)) + p(r('Son not') + endnoteRef(1)) + p(r('Kullanılmayan yok')),
        {
            footnotes: { 1: p(r('Kaynak: ') + hyperlink('rIdNote', r('karar metni'))), 2: p(hyperlink('rIdNote', r('kullanılmayan'))) },
            endnotes: { 1: p(field(`HYPERLINK "${KARAR}"`, r('son karar'))) },
            links: { 'footnotes.xml': { rIdNote: 'https://example.org/kaynak' } }
        });
    assert.deepEqual(occurrences(parsed, 'hyperlinks'), [
        { example: 'karar metni → example.org/kaynak', location: { kind: 'paragraph', paragraph: 1, note: { kind: 'footnote' } } },
        { example: 'karar metni → example.org/kaynak', location: { kind: 'paragraph', paragraph: 2, note: { kind: 'footnote' } } },
        { example: 'son karar → example.org/karar', location: { kind: 'paragraph', paragraph: 3, note: { kind: 'endnote' } } }
    ]);
    assert.equal(finding(parsed, 'hyperlinks').count, 3);
});

test('a link in a cell the generator writes empty is not reported', async () => {
    // The second row's cell continues the first row's vertical merge, so its content isn't converted
    const merged = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p(r('Üst'))}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${p(hyperlink('rIdLink1', r('gizli bağlantı')))}</w:tc></w:tr></w:tbl>`;
    const { parsed } = await convertBody(merged + p(r('Son')), { links: { 'document.xml': { rIdLink1: KARAR } } });
    assert.equal(finding(parsed, 'hyperlinks'), undefined);
});

test('the note only reports: the .udf is the same as for the document with the links\' text as plain runs', async () => {
    const linked = await convert(await buildDocx(
        p(r('Karar için bkz. ') + hyperlink('rIdLink1', r('Yargıtay kararı'))) +
        p(r('Field: ') + field(`HYPERLINK "${KARAR}"`, r('buradan'))) +
        p(r('e-Devlet: ') + fldSimple(`HYPERLINK &quot;${KARAR}&quot;`, r('buradan erişilebilir'))),
        { links: { 'document.xml': { rIdLink1: KARAR } } }));
    const plainText = await convert(await buildDocx(
        p(r('Karar için bkz. ') + r('Yargıtay kararı')) +
        p(r('Field: ') + field('REF Karar', r('buradan'))) +
        p(r('e-Devlet: ') + fldSimple('REF Karar', r('buradan erişilebilir')))));

    assert.equal(linked.xml, plainText.xml);
    assert.equal(finding(linked.parsed, 'hyperlinks').count, 3);
    assert.equal(finding(plainText.parsed, 'hyperlinks'), undefined);
});

test('footnotes with a malformed relationships part are still converted', async () => {
    const docx = await buildDocx(p(r('Dipnot') + footnoteRef(1)),
        { footnotes: { 1: p(r('Kaynak: ') + hyperlink('rIdNote', r('karar metni'))) }, links: { 'footnotes.xml': { rIdNote: KARAR } } });
    const zip = await JSZip.loadAsync(docx);
    zip.file('word/_rels/footnotes.xml.rels', '<Relationships><Relationship');
    const { content, parsed } = await convert(await zip.generateAsync({ type: 'nodebuffer' }));
    assert.equal(content.includes('Kaynak: karar metni'), true);
    assert.equal(finding(parsed, 'notesMoved').count, 1);
});
