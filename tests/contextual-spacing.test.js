const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

// Contextual spacing (w:contextualSpacing, ECMA-376 17.3.1.9): a paragraph that has it drops its
// space before or after where the paragraph next to it has the same paragraph style. Word's
// List Paragraph style sets it, so list items sit without a gap between them.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const pStyle = id => `<w:pStyle w:val="${id}"/>`;
const styled = (id, text, pPr = '') => p(r(text), pStyle(id) + pPr);
const CONTEXTUAL = '<w:contextualSpacing/>';
const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
const REV = 'w:author="Avukat" w:date="2026-01-01T00:00:00Z"';
const deletedParagraph = (id, text) => p(`<w:del w:id="2" ${REV}><w:r><w:delText>${text}</w:delText></w:r></w:del>`,
    `${pStyle(id)}<w:rPr><w:del w:id="5" ${REV}/></w:rPr>`);

const style = (id, pPr, { basedOn = 'Normal', isDefault = false } = {}) =>
    `<w:style w:type="paragraph"${isDefault ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${id}"/>` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : '') + `<w:pPr>${pPr}</w:pPr></w:style>`;
// Word's default template: Normal has 8 pt after, List Paragraph is based on it and has contextual spacing
const NORMAL = style('Normal', '<w:spacing w:after="160"/>', { basedOn: null, isDefault: true });
const LIST_PARAGRAPH = style('ListParagraph', `<w:ind w:left="720"/>${CONTEXTUAL}`);
const STYLES = NORMAL + LIST_PARAGRAPH;

const tbl = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    rows.map(cells => `<w:tr>${cells.map(cell => `<w:tc>${cell}</w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>';
// A floating text box the way Word writes it, DrawingML with a VML fallback
const textBox = inner => '<w:r><mc:AlternateContent>' +
    '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';
const part = (kind, inner) => ({
    name: `${kind}1.xml`, root: kind === 'header' ? 'hdr' : 'ftr', type: kind,
    id: `rId${kind === 'header' ? 'Header' : 'Footer'}1`, inner
});
const REFERENCES = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';

/** The UDF paragraphs holding this text, one per content element */
function paragraphsOf(doc, content, text) {
    return [...doc.querySelectorAll('content[startOffset]')]
        .filter(el => textOf(el, content) === text)
        .map(el => el.closest('paragraph'));
}

/** [SpaceAbove, SpaceBelow] of the one UDF paragraph holding this text (null when absent) */
function spacingOf(doc, content, text) {
    const found = paragraphsOf(doc, content, text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} paragraphs`);
    return [found[0].getAttribute('SpaceAbove'), found[0].getAttribute('SpaceBelow')];
}

/** [SpaceAbove, SpaceBelow] of each text's paragraph */
const spacings = (result, texts) => texts.map(text => spacingOf(result.doc, result.content, text));

test('List Paragraph items have no gap between them, but keep the space after the last item', async () => {
    const texts = ['Giriş', 'Birinci', 'İkinci', 'Üçüncü', 'Sonuç'];
    const result = await convertBody(
        p(r('Giriş')) + styled('ListParagraph', 'Birinci') + styled('ListParagraph', 'İkinci') +
        styled('ListParagraph', 'Üçüncü') + p(r('Sonuç'), pStyle('Normal')), { styles: STYLES });

    assert.deepEqual(spacings(result, texts), [
        [null, '8.0'],
        [null, null],
        [null, null],
        [null, '8.0'],
        [null, '8.0']
    ]);
    assert.deepEqual(findingSummaries(result.parsed.findings), []);
});

test('only the first item keeps the list style\'s space before', async () => {
    const styles = NORMAL + style('ListParagraph', `<w:spacing w:before="120"/>${CONTEXTUAL}`);
    const result = await convertBody(
        p(r('Giriş')) + styled('ListParagraph', 'Birinci') + styled('ListParagraph', 'İkinci') +
        styled('ListParagraph', 'Üçüncü') + p(r('Sonuç')), { styles });

    assert.deepEqual(spacings(result, ['Birinci', 'İkinci', 'Üçüncü']), [
        ['6.0', null],
        [null, null],
        [null, '8.0']
    ]);
});

test('the paragraph\'s own w:contextualSpacing turns the style\'s on or off', async () => {
    const styles = NORMAL + style('ListParagraph', `<w:spacing w:before="120"/>${CONTEXTUAL}`) +
        style('Govde', '<w:spacing w:before="120"/>');
    const result = await convertBody(
        p(r('Giriş')) +
        styled('ListParagraph', 'Açık') +
        styled('ListParagraph', 'Kapalı', '<w:contextualSpacing w:val="0"/>') +
        styled('ListParagraph', 'Yanlış', '<w:contextualSpacing w:val="false"/>') +
        styled('ListParagraph', 'Son') +
        styled('Govde', 'Gövde bir', CONTEXTUAL) +
        styled('Govde', 'Gövde iki', CONTEXTUAL), { styles });

    assert.deepEqual(spacings(result, ['Açık', 'Kapalı', 'Yanlış', 'Son', 'Gövde bir', 'Gövde iki']), [
        // Next is the same style, and this one has it on
        ['6.0', null],
        // Turned off: both sides keep their space, though the neighbours have it on
        ['6.0', '8.0'],
        ['6.0', '8.0'],
        // Previous has it off, but this one has it on
        [null, '8.0'],
        // A style without it, turned on by the paragraphs themselves
        ['6.0', null],
        [null, '8.0']
    ]);
});

test('only the paragraph with contextual spacing loses its space', async () => {
    const styles = NORMAL + style('Govde', '<w:spacing w:before="120" w:after="160"/>');
    const result = await convertBody(
        styled('Govde', 'İşaretli', CONTEXTUAL) + styled('Govde', 'İşaretsiz') +
        styled('Govde', 'Yine işaretsiz') + styled('Govde', 'Sonra işaretli', CONTEXTUAL), { styles });

    assert.deepEqual(spacings(result, ['İşaretli', 'İşaretsiz', 'Yine işaretsiz', 'Sonra işaretli']), [
        ['6.0', null],
        ['6.0', '8.0'],
        ['6.0', '8.0'],
        [null, '8.0']
    ]);
});

test('paragraphs of different styles keep their space; no w:pStyle is the default style', async () => {
    const styles = STYLES + style('Madde', CONTEXTUAL);
    const result = await convertBody(
        styled('ListParagraph', 'Liste') + styled('Madde', 'Madde') +
        p(r('Stilsiz'), CONTEXTUAL) + styled('Normal', 'Normal', CONTEXTUAL) +
        styled('Bilinmeyen', 'Bilinmeyen stil', CONTEXTUAL) + styled('ListParagraph', 'Son'), { styles });

    assert.deepEqual(spacings(result, ['Liste', 'Madde', 'Stilsiz', 'Normal', 'Bilinmeyen stil', 'Son']), [
        [null, '8.0'],
        [null, '8.0'],
        // No w:pStyle, w:pStyle="Normal" and an unknown style are all the default paragraph style
        [null, null],
        [null, null],
        [null, '8.0'],
        [null, '8.0']
    ]);
});

test('document defaults can turn contextual spacing on for every paragraph', async () => {
    const styles = '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160"/>' +
        `${CONTEXTUAL}</w:pPr></w:pPrDefault></w:docDefaults>` + style('Normal', '', { basedOn: null, isDefault: true });
    const result = await convertBody(p(r('Bir')) + p(r('İki')) + p(r('Üç')), { styles });

    assert.deepEqual(spacings(result, ['Bir', 'İki', 'Üç']), [[null, null], [null, null], [null, '8.0']]);
});

test('it applies inside a table cell and a nested table\'s cell, but not across a table or between cells', async () => {
    const items = (prefix, n) => Array.from({ length: n }, (_, i) => styled('ListParagraph', `${prefix} ${i + 1}`)).join('');
    const result = await convertBody(
        items('Önce', 2) +
        tbl([
            [items('Hücre', 2), items('Komşu', 1)],
            [items('Alt', 1) + tbl([[items('İç', 2)]]) + items('Sonra', 1), items('Tek', 1)]
        ]) +
        items('Ardından', 1), { styles: STYLES });

    assert.deepEqual(spacings(result, ['Önce 1', 'Önce 2', 'Hücre 1', 'Hücre 2', 'Komşu 1', 'Alt 1', 'İç 1', 'İç 2',
        'Sonra 1', 'Tek 1', 'Ardından 1']).map(([, below]) => below), [
        null,
        // The table follows
        '8.0',
        null,
        // The next cell's paragraph isn't adjacent
        '8.0',
        '8.0',
        // A nested table follows
        '8.0',
        null,
        '8.0',
        '8.0',
        '8.0',
        '8.0'
    ]);
});

test('it applies in the converted header and footer, but not across a table there', async () => {
    const items = (prefix, n) => Array.from({ length: n }, (_, i) => styled('ListParagraph', `${prefix} ${i + 1}`)).join('');
    const result = await convertBody(p(r('Gövde')), {
        styles: STYLES,
        sectPr: REFERENCES,
        parts: [
            part('header', items('Üst', 2) + tbl([[p(r('Tablo'))]]) + items('Tablodan sonra', 1)),
            part('footer', items('Alt', 2))
        ]
    });

    assert.deepEqual(spacings(result, ['Üst 1', 'Üst 2', 'Tablodan sonra 1', 'Alt 1', 'Alt 2']).map(([, below]) => below),
        [null, '8.0', '8.0', null, '8.0']);
});

test('a paragraph Word\'s Accept All removes doesn\'t separate the paragraphs around it', async () => {
    const result = await convertBody(
        styled('ListParagraph', 'Birinci') + deletedParagraph('Normal', 'Silinen') + styled('ListParagraph', 'İkinci') +
        tbl([[styled('ListParagraph', 'Hücre bir') + deletedParagraph('Normal', 'Silinen hücre') +
            styled('ListParagraph', 'Hücre iki')]]),
        { styles: STYLES });

    assert.equal(result.content.includes('Silinen'), false);
    assert.deepEqual(spacings(result, ['Birinci', 'İkinci', 'Hücre bir', 'Hücre iki']).map(([, below]) => below),
        [null, '8.0', null, '8.0']);
});

test('it applies across a page break between paragraphs, and to a paragraph a page break splits as a whole', async () => {
    const styles = NORMAL + style('ListParagraph', `<w:spacing w:before="120"/>${CONTEXTUAL}`);
    const result = await convertBody(
        p(r('Giriş')) +
        styled('ListParagraph', 'Birinci') +
        styled('ListParagraph', 'Yeni sayfa', '<w:pageBreakBefore/>') +
        p(r('Bölünen başı') + PAGE_BREAK + r('Bölünen sonu'), pStyle('ListParagraph')) +
        styled('ListParagraph', 'Son'), { styles });

    assert.equal(result.doc.querySelectorAll('page-break').length, 2);
    assert.deepEqual(spacings(result, ['Birinci', 'Yeni sayfa', 'Bölünen başı', 'Bölünen sonu', 'Son']), [
        ['6.0', null],
        [null, null],
        // The split paragraph's previous and next paragraphs have its style: no space on either side
        [null, null],
        [null, null],
        [null, '8.0']
    ]);
});

test('a text box\'s paragraphs are a sequence of their own', async () => {
    const result = await convertBody(
        p(r('Çapa') + textBox(styled('ListParagraph', 'Kutu bir') + styled('ListParagraph', 'Kutu iki')), pStyle('ListParagraph')) +
        styled('ListParagraph', 'Gövde'), { styles: STYLES });

    assert.deepEqual(spacings(result, ['Çapa', 'Kutu bir', 'Kutu iki', 'Gövde']).map(([, below]) => below),
        // The anchor and the next body paragraph are adjacent; the box's last paragraph has nothing after it
        [null, null, '8.0', '8.0']);
});

test('without w:contextualSpacing, same-style paragraphs keep their space', async () => {
    const styles = NORMAL + style('ListParagraph', '<w:ind w:left="720"/>');
    const result = await convertBody(styled('ListParagraph', 'Bir') + styled('ListParagraph', 'İki'), { styles });

    assert.deepEqual(spacings(result, ['Bir', 'İki']), [[null, '8.0'], [null, '8.0']]);
});
