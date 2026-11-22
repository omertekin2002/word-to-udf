// Page flow and section layout: a section break that starts the next section on a new page, and a
// paragraph set to start on a new page (Page break before), each write a <page-break>, the way a
// manual page break does; a section layout the single <pageFormat> can't keep is reported
// (sectionLayout).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
const footnoteRef = id => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;

// A section's w:type, which says how that section starts after the previous one
const type = value => `<w:type w:val="${value}"/>`;
// A paragraph that ends a section: the section's w:sectPr in its w:pPr
const sectionEnd = (text, inner = '', pPr = '') => p(r(text), `${pPr}<w:sectPr>${inner}</w:sectPr>`);

const A4_PORTRAIT = '<w:pgSz w:w="11906" w:h="16838"/>';
const A4_LANDSCAPE = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>';
const margins = (top = 1440, right = 1440, bottom = 1440, left = 1440) =>
    `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="708" w:footer="708" w:gutter="0"/>`;
const PORTRAIT_SETUP = A4_PORTRAIT + margins();
const columns = (num, inner = '') => `<w:cols w:num="${num}" w:space="708">${inner}</w:cols>`;

const PAGE_BREAK_BEFORE = '<w:pageBreakBefore/>';
const paragraphStyle = (id, pPr, basedOn = '') => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>` +
    `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}<w:pPr>${pPr}</w:pPr></w:style>`;
const pStyle = id => `<w:pStyle w:val="${id}"/>`;

/** The body's elements in order: 'page-break', or a paragraph's or table's text */
function order({ doc, content }) {
    return [...doc.querySelector('elements').children].map(el => el.localName === 'page-break'
        ? 'page-break'
        : [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));
}

const sectionLayout = parsed => findingSummaries(parsed.findings).find(finding => finding.code === 'sectionLayout');
const occurrences = parsed => JSON.parse(JSON.stringify(parsed.findings))
    .filter(finding => finding.code === 'sectionLayout').flatMap(finding => finding.occurrences);

// Section breaks

test('a section break whose next section starts on a new page writes a page break', async () => {
    for (const sectPr of [type('nextPage') + PORTRAIT_SETUP, PORTRAIT_SETUP, undefined]) {
        const result = await convertBody(p(r('Before')) + sectionEnd('End', PORTRAIT_SETUP) + p(r('After')), { sectPr });
        assert.deepEqual(order(result), ['Before', 'End', 'page-break', 'After'], `final sectPr ${sectPr}`);
        assert.equal(sectionLayout(result.parsed), undefined);
    }
});

test('a continuous or next-column section break writes no page break', async () => {
    for (const start of ['continuous', 'nextColumn']) {
        const result = await convertBody(p(r('Before')) + sectionEnd('End', PORTRAIT_SETUP) + p(r('After')),
            { sectPr: type(start) + PORTRAIT_SETUP });
        assert.deepEqual(order(result), ['Before', 'End', 'After'], start);
    }
});

test('the break type is the w:type of the section that starts after the break, not of the one it ends', async () => {
    // The first section's own w:type says how it starts, which is the start of the document
    const continuousFirst = await convertBody(sectionEnd('End', type('continuous') + PORTRAIT_SETUP) + p(r('After')),
        { sectPr: type('nextPage') + PORTRAIT_SETUP });
    assert.deepEqual(order(continuousFirst), ['End', 'page-break', 'After']);

    // Three sections: the second break's type is in the final w:sectPr, the first break's in the second section's
    const three = await convertBody(sectionEnd('One', PORTRAIT_SETUP) + sectionEnd('Two', type('continuous') + PORTRAIT_SETUP) +
        p(r('Three')), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(three), ['One', 'Two', 'page-break', 'Three']);
});

test('a section break that ends the document writes no trailing page break', async () => {
    const result = await convertBody(p(r('Before')) + sectionEnd('End', PORTRAIT_SETUP), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(result), ['Before', 'End']);
});

test('a section break before a table breaks the page before the table', async () => {
    const table = `<w:tbl><w:tr><w:tc>${p(r('Hücre'))}</w:tc></w:tr></w:tbl>`;
    const result = await convertBody(sectionEnd('End', PORTRAIT_SETUP) + table, { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(result), ['End', 'page-break', 'Hücre']);
});

test('a section break in a block-level content control is still a section break', async () => {
    const result = await convertBody(`<w:sdt><w:sdtContent>${sectionEnd('End', PORTRAIT_SETUP)}</w:sdtContent></w:sdt>` +
        p(r('After')), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(result), ['End', 'page-break', 'After']);
});

// Page break before

test('a paragraph with Page break before starts on a new page', async () => {
    const result = await convertBody(p(r('Önce')) + p(r('EKLER'), PAGE_BREAK_BEFORE) + p(r('Sonra')));
    assert.deepEqual(order(result), ['Önce', 'page-break', 'EKLER', 'Sonra']);
});

test('Page break before comes from the paragraph style and its w:basedOn chain, and the paragraph can turn it off', async () => {
    const styles = paragraphStyle('Ekler', PAGE_BREAK_BEFORE) + paragraphStyle('EklerAlt', '<w:jc w:val="center"/>', 'Ekler');
    const result = await convertBody(p(r('Önce')) + p(r('Stil'), pStyle('Ekler')) + p(r('Miras'), pStyle('EklerAlt')) +
        p(r('Kapalı'), pStyle('Ekler') + '<w:pageBreakBefore w:val="0"/>') +
        p(r('Yanlış'), pStyle('Ekler') + '<w:pageBreakBefore w:val="false"/>'), { styles });
    assert.deepEqual(order(result), ['Önce', 'page-break', 'Stil', 'page-break', 'Miras', 'Kapalı', 'Yanlış']);
});

test('the body\'s first paragraph never gets a page break before it', async () => {
    const result = await convertBody(p(r('İlk'), PAGE_BREAK_BEFORE) + p(r('İkinci')));
    assert.deepEqual(order(result), ['İlk', 'İkinci']);
});

test('a paragraph in a table cell or a text box never gets a page break before it', async () => {
    const table = `<w:tbl><w:tr><w:tc>${p(r('Hücre bir')) + p(r('Hücre iki'), PAGE_BREAK_BEFORE)}</w:tc></w:tr></w:tbl>`;
    const textBox = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="1" cy="1"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx>' +
        `<w:txbxContent>${p(r('Kutu'), PAGE_BREAK_BEFORE)}</w:txbxContent>` +
        '</wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>';
    const result = await convertBody(p(r('Önce')) + table + p(r('Çapa') + textBox));
    assert.deepEqual(order(result), ['Önce', 'Hücre birHücre iki', 'Çapa', 'Kutu']);
});

// No doubled breaks

test('a manual page break followed by a section break or a page-break-before paragraph writes one page break', async () => {
    const sectionAfter = await convertBody(p(r('Önce')) + p(PAGE_BREAK, `<w:sectPr>${PORTRAIT_SETUP}</w:sectPr>`) + p(r('Sonra')),
        { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(sectionAfter), ['Önce', 'page-break', 'Sonra']);

    const breakBeforeAfter = await convertBody(p(r('Önce')) + p(PAGE_BREAK) + p(r('EKLER'), PAGE_BREAK_BEFORE));
    assert.deepEqual(order(breakBeforeAfter), ['Önce', 'page-break', 'EKLER']);

    const textThenBreak = await convertBody(p(r('Önce') + PAGE_BREAK) + p(r('EKLER'), PAGE_BREAK_BEFORE));
    assert.deepEqual(order(textThenBreak), ['Önce', 'page-break', 'EKLER']);
});

test('a section break followed by a page-break-before paragraph writes one page break', async () => {
    const result = await convertBody(sectionEnd('End', PORTRAIT_SETUP) + p(r('EKLER'), PAGE_BREAK_BEFORE), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(result), ['End', 'page-break', 'EKLER']);
});

test('a footnote referenced before a section break is written before its page break', async () => {
    const result = await convertBody(p(r('Metin') + footnoteRef(1)) + sectionEnd('End', PORTRAIT_SETUP) + p(r('Sonra')), {
        sectPr: PORTRAIT_SETUP,
        footnotes: { 1: p(r('Dipnot')) }
    });
    const children = order(result);
    assert.equal(children[0], 'Metin1');
    assert.equal(children[1], 'End');
    const pageBreak = children.indexOf('page-break');
    assert.ok(children.slice(2, pageBreak).some(text => text.includes('Dipnot')), 'the footnote comes before the page break');
    assert.deepEqual(children.slice(pageBreak + 1), ['Sonra']);
    assert.equal(children.filter(child => child === 'page-break').length, 1);
});

// sectionLayout

test('an earlier landscape section in a portrait document is reported where its section ends', async () => {
    const result = await convertBody(p(r('Dilekçe')) + sectionEnd('Ek tablo', A4_LANDSCAPE + margins()) + p(r('Son')),
        { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(sectionLayout(result.parsed), { code: 'sectionLayout', count: 1, examples: ['Ek tablo'] });
    assert.deepEqual(occurrences(result.parsed), [{ example: 'Ek tablo', location: { kind: 'paragraph', paragraph: 2 } }]);
    assert.deepEqual(order(result), ['Dilekçe', 'Ek tablo', 'page-break', 'Son']);
});

test('an earlier section with different margins is reported; the same margins written in other units are not', async () => {
    const different = await convertBody(sectionEnd('Bölüm', A4_PORTRAIT + margins(1440, 1440, 1440, 2268)) + p(r('Son')),
        { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(sectionLayout(different.parsed), { code: 'sectionLayout', count: 1, examples: ['Bölüm'] });

    const same = await convertBody(sectionEnd('Bölüm',
        '<w:pgSz w:w="595.3pt" w:h="841.9pt"/><w:pgMar w:top="1in" w:right="2.54cm" w:bottom="72pt" w:left="25.4mm"/>') +
        p(r('Son')), { sectPr: PORTRAIT_SETUP });
    assert.equal(sectionLayout(same.parsed), undefined);
});

test('a section with more than one column is reported, the final one for the whole document', async () => {
    const earlier = await convertBody(sectionEnd('İki sütun', PORTRAIT_SETUP + columns(2)) + p(r('Son')), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(occurrences(earlier.parsed), [{ example: 'İki sütun', location: { kind: 'paragraph', paragraph: 1 } }]);

    const colChildren = await convertBody(sectionEnd('Sütunlar', PORTRAIT_SETUP +
        '<w:cols w:equalWidth="0"><w:col w:w="4000" w:space="708"/><w:col w:w="4000"/></w:cols>') + p(r('Son')),
    { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(sectionLayout(colChildren.parsed), { code: 'sectionLayout', count: 1, examples: ['Sütunlar'] });

    const final = await convertBody(p(r('Metin')), { sectPr: PORTRAIT_SETUP + columns(2) });
    assert.deepEqual(occurrences(final.parsed), [{ example: '', location: { kind: 'document' } }]);
});

test('an odd- or even-page section break is written as a plain page break and reported', async () => {
    for (const start of ['oddPage', 'evenPage']) {
        const result = await convertBody(p(r('Dilekçe')) + sectionEnd('Bitiş', PORTRAIT_SETUP) + p(r('EKLER')),
            { sectPr: type(start) + PORTRAIT_SETUP });
        assert.deepEqual(order(result), ['Dilekçe', 'Bitiş', 'page-break', 'EKLER'], start);
        assert.deepEqual(occurrences(result.parsed), [{ example: 'Bitiş', location: { kind: 'paragraph', paragraph: 2 } }], start);
    }
});

test('sectionLayout is counted once per section', async () => {
    const result = await convertBody(sectionEnd('Bir', A4_LANDSCAPE + margins(720) + columns(2)) + sectionEnd('İki', A4_LANDSCAPE) +
        p(r('Son')), { sectPr: type('oddPage') + PORTRAIT_SETUP + columns(3) });
    assert.deepEqual(sectionLayout(result.parsed), { code: 'sectionLayout', count: 3, examples: ['Bir', 'İki'] });
});

test('sections that share the final section\'s setup, or a single one-column section, report nothing', async () => {
    const shared = await convertBody(sectionEnd('Bir', PORTRAIT_SETUP + columns(1)) + sectionEnd('İki', PORTRAIT_SETUP) +
        p(r('Son')), { sectPr: type('nextPage') + PORTRAIT_SETUP + columns(1) });
    assert.deepEqual(findingSummaries(shared.parsed.findings), []);

    const single = await convertBody(p(r('Metin')), { sectPr: PORTRAIT_SETUP + '<w:cols w:space="708"/>' });
    assert.deepEqual(findingSummaries(single.parsed.findings), []);

    // A section break with no page setup of its own has nothing to compare
    const bare = await convertBody(sectionEnd('Bir') + p(r('Son')), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(findingSummaries(bare.parsed.findings), []);
});

test('a paragraph left out with its deleted mark writes no section break and reports nothing', async () => {
    const deletedMark = '<w:rPr><w:del w:id="1" w:author="Y"/></w:rPr>';
    const result = await convertBody(p(r('Önce')) +
        p('', `${deletedMark}<w:sectPr>${A4_LANDSCAPE}</w:sectPr>`) + p(r('Sonra')), { sectPr: PORTRAIT_SETUP });
    assert.deepEqual(order(result), ['Önce', 'Sonra']);
    assert.equal(sectionLayout(result.parsed), undefined);
});
