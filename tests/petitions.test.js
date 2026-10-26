const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { JSDOM } = require('jsdom');
const { convert, assertOffsetsInBounds, textOf } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

const [dava, itiraz, tablo] = petitions;
const { DOMParser } = new JSDOM('').window;
const TURKISH = 'ğüşıöçİĞÜŞÖÇ';

// Convert each fixture once and share the result between tests
const conversions = new Map();
function converted(petition) {
    if (!conversions.has(petition)) {
        conversions.set(petition, petition.build().then(convert));
    }
    return conversions.get(petition);
}

function contentElements(doc, root = doc) {
    return [...root.querySelectorAll('content[startOffset]')]
        .filter(el => el.closest('elements'));
}

function elementsWithText(doc, content, text) {
    const found = contentElements(doc).filter(el => textOf(el, content) === text);
    assert.ok(found.length > 0, `no content element holds ${JSON.stringify(text)}`);
    return found;
}

const elementWithText = (doc, content, text) => {
    const found = elementsWithText(doc, content, text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} elements`);
    return found[0];
};

const paragraphOf = (doc, content, text) => elementsWithText(doc, content, text)[0].closest('paragraph');

const occurrences = (text, needle) => text.split(needle).length - 1;

/**
 * Each expected piece must appear in order, and between pieces the content may only
 * hold what the generator adds itself: paragraph and cell separators, empty-paragraph
 * placeholders, tabs and the footnote separator line. So nothing is lost, nothing is
 * repeated and nothing is reordered.
 */
function assertTextInOrder(content, pieces) {
    const gap = /^[\s​─]*$/;
    let cursor = 0;
    for (const piece of pieces) {
        assert.ok(piece.trim(), 'expected pieces must not be blank');
        const at = content.indexOf(piece, cursor);
        assert.ok(at >= 0, `${JSON.stringify(piece)} is missing after offset ${cursor}`);
        assert.match(content.slice(cursor, at), gap,
            `unexpected content before ${JSON.stringify(piece)}: ${JSON.stringify(content.slice(cursor, at))}`);
        cursor = at + piece.length;
    }
    assert.match(content.slice(cursor), gap, `unexpected trailing content: ${JSON.stringify(content.slice(cursor))}`);
}

// What the UDF content should hold: the body, then the footnotes it appends as "N. text"
const expectedPieces = petition =>
    [...petition.texts, ...petition.footnotes.map((text, i) => `${i + 1}. ${text}`)];

const listParagraph = (doc, content, item) => {
    const para = paragraphOf(doc, content, item.text);
    assert.ok(para, `${JSON.stringify(item.text)} is not in a paragraph`);
    return para;
};

for (const petition of petitions) {
    // The parser turns a malformed document.xml into an empty document, so check the input first
    test(`${petition.name}: the .docx is well-formed and wired like a Word-saved file`, async () => {
        const zip = await JSZip.loadAsync(await petition.build());
        const parts = {};
        for (const name of Object.keys(zip.files).filter(name => /\.(xml|rels)$/.test(name))) {
            parts[name] = new DOMParser().parseFromString(await zip.file(name).async('string'), 'application/xml');
            assert.equal(parts[name].getElementsByTagName('parsererror').length, 0, `${name} must be well-formed`);
        }

        const overrides = [...parts['[Content_Types].xml'].getElementsByTagName('Override')]
            .map(el => [el.getAttribute('PartName'), el.getAttribute('ContentType')]);
        const rels = [...parts['word/_rels/document.xml.rels'].getElementsByTagName('Relationship')]
            .map(el => [el.getAttribute('Type').split('/').pop(), el.getAttribute('Target')]);
        const optional = ['styles', 'numbering', 'footnotes'].filter(type => type in petition.options);
        for (const type of optional) {
            assert.ok(zip.file(`word/${type}.xml`), `word/${type}.xml`);
            assert.ok(overrides.some(([name, contentType]) => name === `/word/${type}.xml` &&
                contentType === `application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml`),
            `content type for ${type}.xml`);
            assert.ok(rels.some(([relType, target]) => relType === type && target === `${type}.xml`), `relationship to ${type}.xml`);
        }

        const body = parts['word/document.xml'].getElementsByTagName('w:body')[0];
        assert.equal(body.lastElementChild.tagName, 'w:sectPr', 'the body ends with its section properties');
    });

    test(`${petition.name}: content.xml is well-formed and every offset is in bounds`, async () => {
        // convert() asserts well-formedness
        assertOffsetsInBounds(await converted(petition));
    });

    test(`${petition.name}: every text appears exactly once, in document order`, async () => {
        const { content } = await converted(petition);
        assertTextInOrder(content, expectedPieces(petition));
        for (const text of petition.deleted) {
            assert.equal(occurrences(content, text.trim()), 0, `deleted ${JSON.stringify(text)} must not appear`);
        }
    });

    test(`${petition.name}: Turkish characters survive unchanged`, async () => {
        const { content, xml } = await converted(petition);
        const expected = expectedPieces(petition).join('');
        for (const char of TURKISH) {
            assert.equal(occurrences(content, char), occurrences(expected, char), `count of ${char}`);
        }
        // Still precomposed (İ is U+0130, not I + combining dot) and declared as UTF-8
        assert.equal(content, content.normalize('NFC'));
        assert.match(xml, /^<\?xml version="1.0" encoding="UTF-8"/);
    });

    test(`${petition.name}: list paragraphs carry their list level`, async () => {
        const { doc, content } = await converted(petition);
        for (const item of petition.listItems) {
            const para = listParagraph(doc, content, item);
            assert.equal(para.getAttribute('ListLevel'), String(item.level), item.text);
            assert.ok(para.getAttribute('ListId'), `${item.text} has no ListId`);
        }
    });
}

test('the corpus uses every Turkish letter', () => {
    const all = petitions.flatMap(expectedPieces).join('');
    for (const char of TURKISH) assert.ok(all.includes(char), `no fixture uses ${char}`);
});

test('dava dilekçesi: heading, party block, footnotes and signature keep their layout', async () => {
    const { doc, content } = await converted(dava);

    const heading = elementWithText(doc, content, dava.heading);
    assert.equal(heading.getAttribute('bold'), 'true');
    assert.equal(heading.closest('paragraph').getAttribute('Alignment'), '1');

    for (const label of ['DAVACI', 'VEKİLİ', 'DAVALI', 'KONU']) {
        const el = elementWithText(doc, content, label);
        assert.equal(el.getAttribute('bold'), 'true', label);
        assert.equal(el.closest('paragraph').querySelectorAll('tab').length, 2, `${label} line keeps its tabs`);
    }

    for (const number of ['1', '2']) {
        assert.equal(elementWithText(doc, content, number).getAttribute('superscript'), 'true',
            `footnote reference ${number}`);
    }

    const conclusion = dava.texts.find(text => text.startsWith('Yukarıda açıklanan'));
    assert.equal(paragraphOf(doc, content, conclusion).getAttribute('Alignment'), '3');

    const signature = ['Davacı Vekili', 'Av. Ad Soyad', '(e-imzalıdır)']
        .map(text => elementsWithText(doc, content, text).at(-1));
    for (const el of signature) assert.equal(el.closest('paragraph').getAttribute('Alignment'), '2');
    assert.equal(signature[2].getAttribute('italic'), 'true');
});

test('icra/itiraz dilekçesi: list levels, page break and justified text', async () => {
    const { doc, content } = await converted(itiraz);

    const top = [...doc.querySelector('elements').children];
    const pageBreaks = top.filter(el => el.localName === 'page-break');
    assert.equal(pageBreaks.length, 1);
    const at = top.indexOf(pageBreaks[0]);
    const lastItem = itiraz.listItems.at(-1).text;
    assert.equal(textOf(contentElements(doc, top[at - 1])[0], content), lastItem, 'the break follows the list');
    assert.equal(textOf(contentElements(doc, top[at + 1])[0], content), 'SONUÇ VE İSTEM', 'the conclusion starts the next page');

    assert.equal(paragraphOf(doc, content, itiraz.headings[0]).getAttribute('Alignment'), '1');
    const reasons = itiraz.texts.find(text => text.startsWith('Açıklanan nedenlerle'));
    assert.equal(paragraphOf(doc, content, reasons).getAttribute('Alignment'), '3');
    assert.equal(paragraphOf(doc, content, 'Ad Soyad').getAttribute('Alignment'), '2');
});

test('ekli tablo dilekçesi: table shape, merged total row and in-cell line break', async () => {
    const { doc, content } = await converted(tablo);

    const tables = doc.querySelectorAll('elements table');
    assert.equal(tables.length, 1);
    assert.equal(tables[0].getAttribute('columnCount'), '4');
    const rows = [...tables[0].querySelectorAll('row')];
    assert.deepEqual(rows.map(row => row.querySelectorAll('cell').length), tablo.tableShape);

    const total = elementWithText(doc, content, 'TOPLAM').closest('cell');
    assert.equal(total.getAttribute('colspan'), '3');
    assert.equal(elementWithText(doc, content, 'TOPLAM').getAttribute('bold'), 'true');

    for (const header of ['Sıra', 'Alacak Türü', 'Açıklama', 'Tutar (TL)']) {
        assert.equal(elementWithText(doc, content, header).getAttribute('bold'), 'true', header);
    }

    const invoiceCell = elementWithText(doc, content, 'Örnek fatura no. 0001').closest('cell');
    assert.deepEqual(contentElements(doc, invoiceCell).map(el => textOf(el, content)),
        ['Örnek fatura no. 0001', '\n', '(00.00.2026 tarihli)']);
    assert.equal(paragraphOf(doc, content, '10.000,00').getAttribute('Alignment'), '2');

    assert.equal(elementWithText(doc, content, 'kalem kalem').getAttribute('underline'), 'true');
    const italic = elementWithText(doc, content, 'faiz hesabına ilişkin haklarımız saklı kalmak kaydıyla');
    assert.equal(italic.getAttribute('italic'), 'true');
    assert.equal(italic.getAttribute('underline'), null);
    assert.equal(italic.closest('paragraph').getAttribute('Alignment'), '3');
});

// Known gaps: recorded as todo so they show in the output as the fidelity to-do list
// without failing the run. Each one should pass once its roadmap item (vision #5) lands.

test('lists are bulleted or numbered as numbering.xml defines them', { todo: 'roadmap #6: read numbering.xml' }, async () => {
    for (const petition of petitions) {
        const { doc, content } = await converted(petition);
        for (const item of petition.listItems) {
            const para = listParagraph(doc, content, item);
            const attribute = item.kind === 'bullet' ? 'Bulleted' : 'Numbered';
            assert.equal(para.getAttribute(attribute), 'true', `${petition.name}: ${item.text} should be ${attribute}`);
        }
    }
});

test('headings take bold and size from their styles.xml style', { todo: 'roadmap #7: resolve styles.xml' }, async () => {
    const { doc, content } = await converted(itiraz);
    for (const heading of itiraz.headings) {
        const el = elementWithText(doc, content, heading);
        assert.equal(el.getAttribute('bold'), 'true', heading);
        assert.equal(el.getAttribute('size'), '14', heading);
    }
});

test('a vertically merged cell spans its rows', { todo: 'roadmap #8: emit rowspan for vMerge' }, async () => {
    const { doc, content } = await converted(tablo);
    const merged = elementWithText(doc, content, 'Fatura alacağı').closest('cell');
    assert.equal(merged.getAttribute('rowspan'), '2');
});

test('page margins come from w:sectPr', { todo: 'roadmap #9: read page setup from w:sectPr' }, async () => {
    for (const petition of petitions) {
        const { doc } = await converted(petition);
        const page = doc.querySelector('properties pageFormat');
        for (const side of ['top', 'right', 'bottom', 'left']) {
            const actual = Number(page.getAttribute(`${side}Margin`));
            const expected = petition.margins[side] / 20;
            assert.ok(Math.abs(actual - expected) < 0.01, `${petition.name}: ${side} margin ${actual}pt, expected ${expected}pt`);
        }
    }
});
