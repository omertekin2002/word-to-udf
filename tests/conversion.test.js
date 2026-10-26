const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr}${runs}</w:p>`;

function contentElements(doc, root = doc) {
    return [...root.querySelectorAll('content[startOffset]')]
        .filter(el => el.closest('elements'));
}

test('tab and text in the same run are both kept', async () => {
    const { doc, content, parsed } = await convertBody(p('<w:r><w:tab/><w:t>Hello</w:t></w:r>'));

    assert.ok(content.includes('\tHello'), JSON.stringify(content));
    // Spread to copy the vm-realm array into this realm for deepEqual
    assert.deepEqual([...parsed.elements[0].runs].map(r => r.type), ['tab', 'text']);

    const tab = doc.querySelector('elements tab');
    assert.ok(tab, 'expected a <tab> element');
    assert.equal(textOf(tab, content), '\t');

    const texts = contentElements(doc).map(el => textOf(el, content));
    assert.deepEqual(texts, ['Hello']);
    assert.equal(
        Number(contentElements(doc)[0].getAttribute('startOffset')),
        Number(tab.getAttribute('startOffset')) + 1
    );
});

test('text, tabs and text in one run keep document order', async () => {
    const { content } = await convertBody(p('<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:tab/></w:r>'));
    assert.equal(content, 'a\tb\t');
});

test('a line break between two texts in one run keeps both, in order', async () => {
    const { doc, content } = await convertBody(p('<w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r>'));

    assert.equal(content, 'A\nB');
    assert.deepEqual(contentElements(doc).map(el => textOf(el, content)), ['A', 'B']);
});

test('w:cr is treated as a line break', async () => {
    const { content } = await convertBody(p('<w:r><w:t>A</w:t><w:cr/><w:t>B</w:t></w:r>'));
    assert.equal(content, 'A\nB');
});

test('run formatting applies to every text piece in the run', async () => {
    const { doc, content } = await convertBody(
        p('<w:r><w:rPr><w:b/></w:rPr><w:tab/><w:t>Bold</w:t><w:br/><w:t>Also bold</w:t></w:r>'));

    const els = contentElements(doc);
    assert.deepEqual(els.map(el => textOf(el, content)), ['Bold', 'Also bold']);
    for (const el of els) assert.equal(el.getAttribute('bold'), 'true');
});

test('a page break run that also carries text keeps the text', async () => {
    const { doc, content } = await convertBody(
        p('<w:r><w:t>Before</w:t></w:r><w:r><w:br w:type="page"/><w:t>After</w:t></w:r>'));

    assert.equal(content, 'BeforeAfter');
    const order = [...doc.querySelector('elements').children].map(el => el.localName);
    assert.deepEqual(order, ['paragraph', 'page-break', 'paragraph']);
    assert.deepEqual(contentElements(doc).map(el => textOf(el, content)), ['Before', 'After']);
});

test('text on both sides of a page break in one run is kept', async () => {
    const { doc, content } = await convertBody(
        p('<w:r><w:t>X</w:t><w:br w:type="page"/><w:t>Y</w:t></w:r><w:r><w:t>Z</w:t></w:r>'));

    assert.equal(content, 'XYZ');
    assert.ok(doc.querySelector('elements page-break'));
    assert.deepEqual(contentElements(doc).map(el => textOf(el, content)), ['X', 'Y', 'Z']);
});

test('a list paragraph split by a page break stays one list item', async () => {
    const { doc } = await convertBody(p(
        '<w:r><w:t>X</w:t><w:br w:type="page"/><w:t>Y</w:t></w:r>',
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:firstLine="360"/></w:pPr>'));

    const [before, after] = doc.querySelectorAll('elements paragraph');
    assert.equal(before.getAttribute('Numbered'), 'true');
    assert.equal(before.getAttribute('FirstLineIndent'), '18.0');
    assert.equal(after.getAttribute('Numbered'), null);
    assert.equal(after.getAttribute('ListId'), null);
    assert.equal(after.getAttribute('FirstLineIndent'), null);
    assert.equal(after.getAttribute('LeftIndent'), '36.0');
});

test('a paragraph holding only a page break emits just the page break', async () => {
    const { doc } = await convertBody(
        p('<w:r><w:t>One</w:t></w:r>') + p('<w:r><w:br w:type="page"/></w:r>') + p('<w:r><w:t>Two</w:t></w:r>'));

    const order = [...doc.querySelector('elements').children].map(el => el.localName);
    assert.deepEqual(order, ['paragraph', 'page-break', 'paragraph']);
});

const table = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    rows.map(cells => `<w:tr>${cells.map(c => `<w:tc>${c}</w:tc>`).join('')}</w:tr>`).join('') +
    '</w:tbl>';

test('a line break inside a table cell is preserved in the cell content', async () => {
    const { doc, content } = await convertBody(table([[
        p('<w:r><w:t>Line one</w:t><w:br/><w:t>Line two</w:t></w:r>'),
        p('<w:r><w:t>Other</w:t></w:r>')
    ]]));

    assert.ok(content.includes('Line one\nLine two'), JSON.stringify(content));
    const cell = doc.querySelector('elements cell');
    assert.deepEqual(contentElements(doc, cell).map(el => textOf(el, content)), ['Line one', 'Line two']);
});

test('a page break inside a table cell becomes a line break', async () => {
    const { doc, content } = await convertBody(table([[
        p('<w:r><w:t>Top</w:t><w:br w:type="page"/><w:t>Bottom</w:t></w:r>'),
        p('<w:r><w:t>Other</w:t></w:r>')
    ]]));

    assert.ok(content.includes('Top\nBottom'), JSON.stringify(content));
    assert.equal(doc.querySelector('elements page-break'), null);
});

test('regression: plain text with bold and italic runs', async () => {
    const { doc, content } = await convertBody(p(
        '<w:r><w:t xml:space="preserve">Plain </w:t></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold </w:t></w:r>' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>',
        '<w:pPr><w:jc w:val="center"/></w:pPr>'));

    assert.equal(content, 'Plain bold italic');
    const [plain, bold, italic] = contentElements(doc);
    assert.equal(textOf(plain, content), 'Plain ');
    assert.equal(plain.getAttribute('bold'), null);
    assert.equal(textOf(bold, content), 'bold ');
    assert.equal(bold.getAttribute('bold'), 'true');
    assert.equal(bold.getAttribute('italic'), null);
    assert.equal(textOf(italic, content), 'italic');
    assert.equal(italic.getAttribute('italic'), 'true');
    assert.equal(doc.querySelector('elements paragraph').getAttribute('Alignment'), '1');
});

test('regression: a simple 2x2 table', async () => {
    const cell = text => p(`<w:r><w:t>${text}</w:t></w:r>`);
    const { doc, content } = await convertBody(table([
        [cell('A1'), cell('B1')],
        [cell('A2'), cell('B2')]
    ]));

    const tbl = doc.querySelector('elements table');
    assert.equal(tbl.getAttribute('columnCount'), '2');
    const rows = [...tbl.querySelectorAll('row')];
    assert.equal(rows.length, 2);
    const cells = rows.map(row => [...row.querySelectorAll('cell')]
        .map(c => contentElements(doc, c).map(el => textOf(el, content)).join('')));
    assert.deepEqual(cells, [['A1', 'B1'], ['A2', 'B2']]);
});

test('regression: a footnote reference becomes a superscript number with the text appended', async () => {
    const { doc, content } = await convertBody(
        p('<w:r><w:t>Claim</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'),
        { footnotes: { 1: p('<w:r><w:footnoteRef/></w:r><w:r><w:t>Source text</w:t></w:r>') } });

    const els = contentElements(doc);
    assert.equal(textOf(els[0], content), 'Claim');
    assert.equal(textOf(els[1], content), '1');
    assert.equal(els[1].getAttribute('superscript'), 'true');

    const tail = els.slice(-2).map(el => textOf(el, content));
    assert.deepEqual(tail, ['1. ', 'Source text']);
    assert.ok(content.endsWith('Source text'));
});
