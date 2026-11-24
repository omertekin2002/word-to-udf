const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

// Superscript and subscript runs (w:vertAlign) and paragraph borders (w:pBdr). Superscript is
// written as superscript="true", as the note marks are. Docs.md gives no subscript and borders only
// on tables and cells, so subscript text stays on the baseline and both are reported instead.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const vertAlign = val => `<w:vertAlign w:val="${val}"/>`;
const sup = text => r(text, vertAlign('superscript'));
const sub = text => r(text, vertAlign('subscript'));
const edge = (name, val) => `<w:${name} w:val="${val}" w:sz="4" w:space="1" w:color="auto"/>`;
const pBdr = (...edges) => `<w:pBdr>${edges.join('')}</w:pBdr>`;
const boxed = pBdr(edge('top', 'single'), edge('bottom', 'single'));
const tbl = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    rows.map(row => '<w:tr>' + row.map(content => `<w:tc>${content}</w:tc>`).join('') + '</w:tr>').join('') + '</w:tbl>';
const footnoteRef = (id, rPr = '<w:rStyle w:val="FootnoteReference"/>') =>
    `<w:r><w:rPr>${rPr}</w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const endnoteRef = (id, rPr = '<w:rStyle w:val="EndnoteReference"/>') =>
    `<w:r><w:rPr>${rPr}</w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const NOTE_REFERENCE_STYLES =
    '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>' +
    `<w:rPr>${vertAlign('superscript')}</w:rPr></w:style>` +
    '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/>' +
    `<w:rPr>${vertAlign('superscript')}</w:rPr></w:style>`;
const HEADING_STYLE = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;

// Each <content> piece of the body's paragraphs: its text and its attributes
const pieces = (doc, content, selector = 'elements > paragraph > content') => [...doc.querySelectorAll(selector)]
    .map(el => ({ text: textOf(el, content), attrs: Object.fromEntries([...el.attributes]
        .filter(a => !['startOffset', 'length'].includes(a.name)).map(a => [a.name, a.value])) }));
const pieceOf = (doc, content, text) => pieces(doc, content, 'content').find(piece => piece.text === text);

test('a direct superscript run is its own <content> with superscript="true" and its usual family and size', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Alan 120 m', '<w:rFonts w:ascii="Arial"/><w:sz w:val="24"/>') +
            r('2', `<w:rFonts w:ascii="Arial"/><w:sz w:val="24"/>${vertAlign('superscript')}`) + r(' olup')));
    assert.deepEqual(pieces(doc, content), [
        { text: 'Alan 120 m', attrs: { family: 'Arial', size: '12' } },
        { text: '2', attrs: { family: 'Arial', size: '12', superscript: 'true' } },
        { text: ' olup', attrs: { family: 'Times New Roman', size: '12' } }
    ]);
    assert.equal(finding(parsed, 'subscript'), undefined);
    assert.equal(finding(parsed, 'paragraphBorders'), undefined);
});

test('superscript keeps the run\'s other formatting, and All Caps still applies', async () => {
    const { doc, content } = await convertBody(p(r('1') + r('er', `<w:b/><w:caps/>${vertAlign('superscript')}`)));
    assert.deepEqual(pieceOf(doc, content, 'ER').attrs, { family: 'Times New Roman', size: '12', bold: 'true', superscript: 'true' });
});

test('superscript from a character style, a paragraph style or the document defaults is written', async () => {
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>' +
        '<w:style w:type="character" w:styleId="Ust"><w:name w:val="Ust"/><w:rPr>' + vertAlign('superscript') + '</w:rPr></w:style>' +
        '<w:style w:type="character" w:styleId="UstBase"><w:name w:val="UstBase"/><w:basedOn w:val="Ust"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="UstPara"><w:name w:val="UstPara"/><w:rPr>' + vertAlign('superscript') + '</w:rPr></w:style>';
    const { doc, content } = await convertBody(
        p(r('a') + r('karakter', '<w:rStyle w:val="Ust"/>') + r('zincir', '<w:rStyle w:val="UstBase"/>')) +
        p(r('paragraf'), '<w:pStyle w:val="UstPara"/>'), { styles });
    assert.equal(pieceOf(doc, content, 'a').attrs.superscript, undefined);
    for (const text of ['karakter', 'zincir', 'paragraf']) assert.equal(pieceOf(doc, content, text).attrs.superscript, 'true', text);

    const defaults = '<w:docDefaults><w:rPrDefault><w:rPr>' + vertAlign('superscript') + '</w:rPr></w:rPrDefault></w:docDefaults>';
    const fromDefaults = await convertBody(p(r('varsayılan')), { styles: defaults });
    assert.equal(pieceOf(fromDefaults.doc, fromDefaults.content, 'varsayılan').attrs.superscript, 'true');
});

test('a direct baseline over an inherited superscript writes none, and a later layer can switch to subscript', async () => {
    const styles = '<w:style w:type="character" w:styleId="Ust"><w:name w:val="Ust"/><w:rPr>' + vertAlign('superscript') + '</w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="UstPara"><w:name w:val="UstPara"/><w:rPr>' + vertAlign('superscript') + '</w:rPr></w:style>';
    const { doc, content, parsed } = await convertBody(
        p(r('karakter', `<w:rStyle w:val="Ust"/>${vertAlign('baseline')}`) + r(' paragraf', vertAlign('baseline')) +
            r(' alt', `<w:rStyle w:val="Ust"/>${vertAlign('subscript')}`), '<w:pStyle w:val="UstPara"/>'), { styles });
    assert.deepEqual(pieces(doc, content).map(piece => piece.attrs.superscript), [undefined, undefined, undefined]);
    assert.equal(finding(parsed, 'subscript').count, 1);
});

test('superscript is written in table cells, nested tables, footnote text and the header', async () => {
    const header = { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Büro') + sup('®')) };
    const { doc, content } = await convertBody(
        tbl([[p(r('m') + sup('3')), tbl([[p(r('cm') + sup('2')), p('')]])]]) + p(r('Dipnot') + footnoteRef(1)),
        { footnotes: { 1: p(r('Kaynak') + sup('a')) }, parts: [header], sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>' });
    for (const text of ['3', '2', 'a']) assert.equal(pieceOf(doc, content, text).attrs.superscript, 'true', text);
    const headerPiece = [...doc.querySelectorAll('header content')].find(el => textOf(el, content) === '®');
    assert.equal(headerPiece.getAttribute('superscript'), 'true');
});

test('note marks keep exactly one superscript="true" when their reference style is superscript', async () => {
    const { doc, xml, content } = await convertBody(
        p(r('Dipnot') + footnoteRef(1) + r(' ve son not') + endnoteRef(1)) +
        tbl([[p(r('Hücre') + footnoteRef(2, `<w:rStyle w:val="FootnoteReference"/>${vertAlign('superscript')}`)), p('')]]),
        { styles: NOTE_REFERENCE_STYLES, footnotes: { 1: p(r('Bir')), 2: p(r('İki')) }, endnotes: { 1: p(r('Son')) } });
    assert.doesNotMatch(xml, /superscript="true"[^>]*superscript=/);
    const marks = [...doc.querySelectorAll('content[superscript]')].map(el => textOf(el, content));
    // The three reference marks and the three note numbers
    assert.deepEqual(marks.sort(), ['1', '1. ', '2', '2. ', 'i', 'i. '].sort());
});

test('a subscript run stays on the baseline and is reported once per paragraph, quoting the paragraph', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Formül: H') + sub('2') + r('O ve CO') + sub('2')) + p(r('Düz paragraf')));
    assert.deepEqual(pieces(doc, content).slice(0, 4).map(piece => [piece.text, piece.attrs]), [
        ['Formül: H', { family: 'Times New Roman', size: '12' }],
        ['2', { family: 'Times New Roman', size: '12' }],
        ['O ve CO', { family: 'Times New Roman', size: '12' }],
        ['2', { family: 'Times New Roman', size: '12' }]
    ]);
    assert.match(content, /Formül: H2O ve CO2/);
    assert.deepEqual(occurrences(parsed, 'subscript'),
        [{ example: 'Formül: H2O ve CO2', location: { kind: 'paragraph', paragraph: 1 } }]);
    assert.equal(finding(parsed, 'subscript').count, 1);
});

test('a subscript that is hidden or only whitespace is not reported', async () => {
    const { parsed } = await convertBody(p(r('H') + r('2', `<w:vanish/>${vertAlign('subscript')}`) + r('O') + sub(' ')));
    assert.equal(finding(parsed, 'subscript'), undefined);
});

test('subscript is located in a table cell, with the nearest heading, and in a note\'s text', async () => {
    const { parsed } = await convertBody(
        p(r('DELİLLER'), '<w:pStyle w:val="Heading1"/>') + tbl([[p(r('A')), p(r('H') + sub('2') + r('O'))]]) +
        p(r('Dipnot') + footnoteRef(1)),
        { styles: HEADING_STYLE, footnotes: { 1: p(r('CO') + sub('2') + r(' ölçümü')) } });
    assert.deepEqual(occurrences(parsed, 'subscript'), [
        { example: 'H2O', location: { kind: 'cell', table: 1, row: 1, cell: 2, heading: 'DELİLLER' } },
        { example: 'CO2 ölçümü', location: { kind: 'paragraph', paragraph: 2, heading: 'DELİLLER', note: { kind: 'footnote' } } }
    ]);
});

test('a paragraph with a visible border edge, set directly, is reported once, located and quoted, and no border is written', async () => {
    const { xml, parsed } = await convertBody(
        p(r('Giriş')) + p(r('SONUÇ VE İSTEM'), boxed) + p(r('Alt çizgili'), pBdr(edge('bottom', 'double'))));
    assert.deepEqual(occurrences(parsed, 'paragraphBorders'), [
        { example: 'SONUÇ VE İSTEM', location: { kind: 'paragraph', paragraph: 2 } },
        // An all-capitals paragraph counts as a petition heading for what follows it
        { example: 'Alt çizgili', location: { kind: 'paragraph', paragraph: 3, heading: 'SONUÇ VE İSTEM' } }
    ]);
    assert.equal(finding(parsed, 'paragraphBorders').count, 2);
    assert.doesNotMatch(xml, /<paragraph[^>]*[Bb]order/);
});

test('borders from a paragraph style are reported; edges inherit one by one, and all nil or none edges report nothing', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Kutu"><w:name w:val="Kutu"/>' +
        `<w:pPr>${boxed}</w:pPr></w:style>` +
        '<w:style w:type="paragraph" w:styleId="KutuAlt"><w:name w:val="KutuAlt"/><w:basedOn w:val="Kutu"/>' +
        `<w:pPr>${pBdr(edge('top', 'nil'))}</w:pPr></w:style>` +
        '<w:style w:type="paragraph" w:styleId="Cizgisiz"><w:name w:val="Cizgisiz"/>' +
        `<w:pPr>${pBdr(edge('top', 'none'), edge('bottom', 'nil'))}</w:pPr></w:style>`;
    const { parsed } = await convertBody(
        p(r('Stil kutusu'), '<w:pStyle w:val="Kutu"/>') +
        // Only the top edge is switched off; the bottom one still comes from Kutu
        p(r('Alt kenar kaldı'), '<w:pStyle w:val="KutuAlt"/>') +
        p(r('Kenarlar kapalı'), `<w:pStyle w:val="Kutu"/>${pBdr(edge('top', 'nil'), edge('bottom', 'none'))}`) +
        p(r('Çizgisiz stil'), '<w:pStyle w:val="Cizgisiz"/>') +
        p(r('Boş kenarlık'), pBdr(edge('left', 'none'), edge('right', 'nil'))),
        { styles });
    assert.deepEqual(occurrences(parsed, 'paragraphBorders').map(o => o.example), ['Stil kutusu', 'Alt kenar kaldı']);
});

test('paragraph borders are located in a table cell and in a note\'s text', async () => {
    const { parsed } = await convertBody(
        tbl([[p(r('Kutulu hücre'), boxed), p('')]]) + p(r('Dipnot') + footnoteRef(1)),
        { footnotes: { 1: p(r('Kutulu not'), boxed) } });
    assert.deepEqual(occurrences(parsed, 'paragraphBorders'), [
        { example: 'Kutulu hücre', location: { kind: 'cell', table: 1, row: 1, cell: 1 } },
        { example: 'Kutulu not', location: { kind: 'paragraph', paragraph: 1, note: { kind: 'footnote' } } }
    ]);
});

test('the issue\'s reproduction: m² is superscript, and H₂O and the boxed heading each give one note', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Alan 120 m') + sup('2')) + p(r('H') + sub('2') + r('O')) + p(r('SONUÇ VE İSTEM'), boxed));
    assert.equal(pieces(doc, content).find(piece => piece.text === '2' && piece.attrs.superscript).attrs.superscript, 'true');
    assert.equal(finding(parsed, 'subscript').count, 1);
    assert.equal(finding(parsed, 'paragraphBorders').count, 1);
});

test('the petition fixtures give neither new finding', async () => {
    for (const petition of petitions) {
        const { DocxParser } = require('./helpers').loadConverter();
        const parsed = await new DocxParser().parse(await petition.build());
        const codes = plain(parsed.findings).map(f => f.code);
        assert.ok(!codes.includes('subscript') && !codes.includes('paragraphBorders'), petition.name);
    }
});
