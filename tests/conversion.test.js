const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, loadConverter, buildDocx } = require('./helpers');

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
    // The newline is referenced too, so the element tree keeps the break
    assert.deepEqual(contentElements(doc).map(el => textOf(el, content)), ['A', '\n', 'B']);
});

test('w:cr is treated as a line break', async () => {
    const { content } = await convertBody(p('<w:r><w:t>A</w:t><w:cr/><w:t>B</w:t></w:r>'));
    assert.equal(content, 'A\nB');
});

test('run formatting applies to every text piece in the run', async () => {
    const { doc, content } = await convertBody(
        p('<w:r><w:rPr><w:b/></w:rPr><w:tab/><w:t>Bold</w:t><w:br/><w:t>Also bold</w:t></w:r>'));

    const els = contentElements(doc).filter(el => textOf(el, content) !== '\n');
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

test('a list paragraph that starts with a page break keeps its numbering', async () => {
    const { doc } = await convertBody(p(
        '<w:r><w:br w:type="page"/><w:t>Y</w:t></w:r>',
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:firstLine="360"/></w:pPr>'));

    const order = [...doc.querySelector('elements').children].map(el => el.localName);
    assert.deepEqual(order, ['page-break', 'paragraph']);
    const para = doc.querySelector('elements paragraph');
    assert.equal(para.getAttribute('Numbered'), 'true');
    assert.equal(para.getAttribute('FirstLineIndent'), '18.0');
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
    assert.deepEqual(contentElements(doc, cell).map(el => textOf(el, content)), ['Line one', '\n', 'Line two']);
});

test('a page break inside a table cell becomes a line break', async () => {
    const { doc, content } = await convertBody(table([[
        p('<w:r><w:t>Top</w:t><w:br w:type="page"/><w:t>Bottom</w:t></w:r>'),
        p('<w:r><w:t>Other</w:t></w:r>')
    ]]));

    assert.ok(content.includes('Top\nBottom'), JSON.stringify(content));
    assert.equal(doc.querySelector('elements page-break'), null);
    const cell = doc.querySelector('elements cell');
    assert.deepEqual(contentElements(doc, cell).map(el => textOf(el, content)), ['Top', '\n', 'Bottom']);
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

// A table on a grid of 2000-twip columns; each cell is [text, tcPr], with text null for an empty paragraph
const gridTable = (columns, rows, tblPr = '') => `<w:tbl>${tblPr ? `<w:tblPr>${tblPr}</w:tblPr>` : ''}` +
    `<w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>` +
    rows.map(cells => `<w:tr>${cells.map(([text, tcPr = '']) =>
        `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${text === null ? '<w:p/>' : p(`<w:r><w:t>${text}</w:t></w:r>`)}</w:tc>`
    ).join('')}</w:tr>`).join('') +
    '</w:tbl>';
const RESTART = '<w:vMerge w:val="restart"/>';
const CONTINUE = '<w:vMerge/>';
const NO_BORDERS = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(side => `<w:${side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`).join('')}</w:tblBorders>`;

const cellRows = doc => [...doc.querySelectorAll('elements table row')].map(row => [...row.querySelectorAll('cell')]);
const cellText = (doc, cell, content) => contentElements(doc, cell).map(el => textOf(el, content)).join('');
const cellAttribute = name => cells => cells.map(cell => cell.getAttribute(name));

test('cell shading is written as fillColor, not bgColor', async () => {
    const { doc } = await convertBody(gridTable(2, [[
        ['Red', '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>'],
        ['Auto', '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>']
    ]]));

    const [[red, auto]] = cellRows(doc);
    assert.equal(red.getAttribute('fillColor'), String(0xFFFF0000 | 0));
    assert.equal(auto.getAttribute('fillColor'), null);
    for (const cell of [red, auto]) assert.equal(cell.getAttribute('bgColor'), null);
});

test('vertical alignment is written as align with the documented values', async () => {
    const { doc } = await convertBody(gridTable(4, [[
        ['Center', '<w:vAlign w:val="center"/>'],
        ['Bottom', '<w:vAlign w:val="bottom"/>'],
        ['Top', '<w:vAlign w:val="top"/>'],
        ['Unknown', '<w:vAlign w:val="constructor"/>']
    ]]));

    const [cells] = cellRows(doc);
    assert.deepEqual(cellAttribute('align')(cells), ['vcenter', 'bottom', null, null]);
    assert.deepEqual(cellAttribute('vAlign')(cells), [null, null, null, null]);
});

test('a three-row vertical merge keeps its column and joins its cells with borderSpec', async () => {
    const shaded = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>';
    const { doc, content } = await convertBody(gridTable(2, [
        [['Merged', RESTART + shaded], ['B1']],
        [[null, CONTINUE + shaded], ['B2']],
        [[null, CONTINUE], ['B3']],
        [['A4'], ['B4']]
    ]));

    // Continued cells add nothing to the content
    assert.equal(content, 'MergedB1B2B3A4B4');
    const rows = cellRows(doc);
    assert.deepEqual(rows.map(cells => cells.map(cell => cellText(doc, cell, content))),
        [['Merged', 'B1'], ['', 'B2'], ['', 'B3'], ['A4', 'B4']]);
    assert.deepEqual(rows.map(cellAttribute('borderSpec')),
        [['11', null], ['10', null], ['14', null], [null, null]]);
    // The merged area is shaded like its first cell
    const D9D9D9 = String(0xFFD9D9D9 | 0);
    assert.deepEqual(rows.map(cells => cells[0].getAttribute('fillColor')), [D9D9D9, D9D9D9, D9D9D9, null]);
    for (const cell of [rows[1][0], rows[2][0]]) {
        assert.ok(cell.querySelector('paragraph'), 'a continued cell still holds a paragraph');
    }
});

test('a vertical merge next to a gridSpan cell matches cells by grid column', async () => {
    const { doc, content } = await convertBody(gridTable(3, [
        [['Wide', '<w:gridSpan w:val="2"/>'], ['Tall', RESTART]],
        [['Wide 2', '<w:gridSpan w:val="2"/>'], [null, CONTINUE]],
        // Column 2 starts no cell in the row above, so this continuation joins nothing
        [['A3'], ['Orphan', CONTINUE], ['C3']]
    ]));

    const rows = cellRows(doc);
    for (const cells of rows) {
        assert.equal(cells.reduce((sum, cell) => sum + Number(cell.getAttribute('colspan') || '1'), 0), 3);
    }
    assert.deepEqual(rows.map(cells => cells.map(cell => cellText(doc, cell, content))),
        [['Wide', 'Tall'], ['Wide 2', ''], ['A3', 'Orphan', 'C3']]);
    assert.deepEqual(rows.map(cellAttribute('borderSpec')),
        [[null, '11'], [null, '14'], [null, null, null]]);
});

test('merged cells in a borderNone table get no borderSpec', async () => {
    const { doc, content } = await convertBody(gridTable(2, [
        [['Merged', RESTART], ['B1']],
        [[null, CONTINUE], ['B2']]
    ], NO_BORDERS));

    assert.equal(doc.querySelector('elements table').getAttribute('border'), 'borderNone');
    assert.equal(content, 'MergedB1B2');
    const rows = cellRows(doc);
    assert.equal(rows[1].length, 2);
    assert.deepEqual(rows.map(cellAttribute('borderSpec')), [[null, null], [null, null]]);
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

// "]]>" would end the CDATA section early; the content must still be well-formed and decode exactly
for (const text of ['a ]]> b', 'x]]]>y', ']]>]]>', 'ends with ]]']) {
    const escaped = text.replace(/>/g, '&gt;');

    test(`text containing ${JSON.stringify(text)} keeps content.xml well-formed`, async () => {
        const { doc, content } = await convertBody(p(`<w:r><w:t xml:space="preserve">${escaped}</w:t></w:r>`));
        assert.equal(content, text);
        assert.deepEqual(contentElements(doc).map(el => textOf(el, content)), [text]);
    });
}

test('"]]>" in a table cell and a footnote keeps content.xml well-formed', async () => {
    const { doc, content } = await convertBody(
        table([[p('<w:r><w:t>cell ]]&gt; one</w:t></w:r>'), p('<w:r><w:t>two</w:t></w:r>')]]) +
        p('<w:r><w:t>Claim</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'),
        { footnotes: { 1: p('<w:r><w:footnoteRef/></w:r><w:r><w:t>note ]]&gt; text</w:t></w:r>') } });

    const cell = doc.querySelector('elements cell');
    assert.deepEqual(contentElements(doc, cell).map(el => textOf(el, content)), ['cell ]]> one']);
    assert.equal(textOf(contentElements(doc).at(-1), content), 'note ]]> text');
    assert.ok(content.endsWith('note ]]> text'));
});

// Lists: bullet vs. number and the number style come from word/numbering.xml
const listItem = (text, numId, ilvl = 0) =>
    p(`<w:r><w:t>${text}</w:t></w:r>`, `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`);
const lvl = (ilvl, numFmt, lvlText) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/></w:lvl>`;
const abstractNum = (id, levels) => `<w:abstractNum w:abstractNumId="${id}">${levels}</w:abstractNum>`;
const num = (numId, abstractNumId, overrides = '') =>
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/>${overrides}</w:num>`;

function listParagraphs(doc, content) {
    return Object.fromEntries([...doc.querySelectorAll('elements paragraph')]
        .map(para => [contentElements(doc, para).map(el => textOf(el, content)).join(''), para]));
}

test('an even numId can be numbered and an odd one bulleted', async () => {
    const { doc, content } = await convertBody(listItem('Number', 2) + listItem('Bullet', 1), {
        numbering: abstractNum(0, lvl(0, 'bullet', '')) + abstractNum(1, lvl(0, 'decimal', '%1.')) +
            num(1, 0) + num(2, 1)
    });

    const paras = listParagraphs(doc, content);
    assert.equal(paras.Number.getAttribute('Numbered'), 'true');
    assert.equal(paras.Number.getAttribute('Bulleted'), null);
    assert.equal(paras.Number.getAttribute('ListId'), '2');
    assert.equal(paras.Bullet.getAttribute('Bulleted'), 'true');
    assert.equal(paras.Bullet.getAttribute('BulletType'), 'BULLET_TYPE_ELLIPSE');
    assert.equal(paras.Bullet.getAttribute('Numbered'), null);
    assert.equal(paras.Bullet.getAttribute('ListId'), '1');
});

test('NumberType follows numFmt and the suffix in lvlText', async () => {
    const levels = [
        ['decimal', '%1.'], ['decimal', '%1)'], ['lowerLetter', '%2)'], ['lowerRoman', '%3.'],
        ['upperLetter', '%1.'], ['upperLetter', '%1)'], ['upperRoman', '%1.'], ['upperRoman', '%1)'],
        ['lowerLetter', '%1.'], ['lowerRoman', '%1)'], ['decimal', '%1-'], ['decimal', '(%1)'],
        ['decimal', '%1'], ['lowerLetter', '%1:'], ['ordinal', '%1.'], ['decimalZero', '%1)']
    ];
    const expected = [
        'NUMBER_TYPE_NUMBER_DOT', 'NUMBER_TYPE_NUMBER_PARENTHESIS', 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS',
        'NUMBER_TYPE_ROMAN_SMALL_DOT', 'NUMBER_TYPE_CHAR_BIG_DOT', 'NUMBER_TYPE_CHAR_BIG_PARENTHESIS',
        'NUMBER_TYPE_ROMAN_BIG_DOT', 'NUMBER_TYPE_ROMAN_BIG_PARENTHESIS', 'NUMBER_TYPE_CHAR_SMALL_DOT',
        'NUMBER_TYPE_ROMAN_SMALL_PARENTHESIS', 'NUMBER_TYPE_NUMBER_TRE', 'NUMBER_TYPE_NUMBER_PARENTHESIS',
        'NUMBER_TYPE_NUMBER_DOT', 'NUMBER_TYPE_CHAR_SMALL_DOT', 'NUMBER_TYPE_NUMBER_DOT', 'NUMBER_TYPE_NUMBER_DOT'
    ];
    // One list per case; the level is the one whose placeholder the lvlText uses
    const ilvlOf = lvlText => Number(lvlText.match(/%(\d)/)[1]) - 1;
    const { doc, content } = await convertBody(
        levels.map(([, lvlText], i) => listItem(`Item${i}`, i + 1, ilvlOf(lvlText))).join(''),
        {
            numbering: levels.map(([numFmt, lvlText], i) => abstractNum(i, lvl(ilvlOf(lvlText), numFmt, lvlText))).join('') +
                levels.map((_, i) => num(i + 1, i)).join('')
        });

    const paras = listParagraphs(doc, content);
    levels.forEach(([numFmt, lvlText], i) => {
        const para = paras[`Item${i}`];
        assert.equal(para.getAttribute('Numbered'), 'true', `${numFmt} ${lvlText}`);
        assert.equal(para.getAttribute('NumberType'), expected[i], `${numFmt} ${lvlText}`);
        assert.equal(para.getAttribute('ListLevel'), String(ilvlOf(lvlText)), `${numFmt} ${lvlText}`);
    });
});

test('numId 0 and numFmt none are not list items', async () => {
    const { doc, content } = await convertBody(listItem('Off', 0) + listItem('None', 1), {
        numbering: abstractNum(0, lvl(0, 'none', '')) + num(1, 0)
    });

    const paras = listParagraphs(doc, content);
    for (const text of ['Off', 'None']) {
        for (const attribute of ['Numbered', 'Bulleted', 'NumberType', 'BulletType', 'ListLevel', 'ListId']) {
            assert.equal(paras[text].getAttribute(attribute), null, `${text}: ${attribute}`);
        }
    }
});

test('a w:lvlOverride on the num replaces the level format', async () => {
    const { doc, content } = await convertBody(listItem('Overridden', 1) + listItem('Plain', 2) + listItem('Deeper', 1, 1), {
        numbering: abstractNum(0, lvl(0, 'decimal', '%1.') + lvl(1, 'decimal', '%2.')) +
            num(1, 0, `<w:lvlOverride w:ilvl="0">${lvl(0, 'upperRoman', '%1)')}</w:lvlOverride>`) +
            num(2, 0)
    });

    const paras = listParagraphs(doc, content);
    assert.equal(paras.Overridden.getAttribute('NumberType'), 'NUMBER_TYPE_ROMAN_BIG_PARENTHESIS');
    assert.equal(paras.Plain.getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_DOT');
    assert.equal(paras.Deeper.getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_DOT');
});

test('a w:lvlOverride level without w:numFmt is decimal, not the abstract level format', async () => {
    const { doc, content } = await convertBody(listItem('Overridden', 1), {
        numbering: abstractNum(0, lvl(0, 'bullet', '')) +
            num(1, 0, '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride>')
    });

    const para = listParagraphs(doc, content).Overridden;
    assert.equal(para.getAttribute('Bulleted'), null);
    assert.equal(para.getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_PARENTHESIS');
});

test('a w:lvlOverride with only a startOverride keeps the abstract level', async () => {
    const { doc, content } = await convertBody(listItem('Restarted', 1), {
        numbering: abstractNum(0, lvl(0, 'lowerLetter', '%1)')) +
            num(1, 0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')
    });

    assert.equal(listParagraphs(doc, content).Restarted.getAttribute('NumberType'), 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS');
});

test('list paragraphs still convert without numbering.xml or a matching definition', async () => {
    const body = listItem('Odd', 1) + listItem('Even', 2);

    for (const options of [{}, { numbering: abstractNum(0, lvl(0, 'bullet', '')) + num(7, 0) }]) {
        const { doc, content } = await convertBody(body, options);
        const paras = listParagraphs(doc, content);
        // The numId guess from before numbering.xml was read
        assert.equal(paras.Odd.getAttribute('Numbered'), 'true');
        assert.equal(paras.Odd.getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_TRE');
        assert.equal(paras.Odd.getAttribute('ListId'), '1');
        assert.equal(paras.Even.getAttribute('Bulleted'), 'true');
        assert.equal(paras.Even.getAttribute('ListId'), '2');
    }
});

test('a malformed numbering.xml does not stop the conversion', async () => {
    const { doc, content } = await convertBody(listItem('Item', 1), { numbering: '<w:num w:numId="1">' });
    assert.equal(content, 'Item');
    assert.equal(listParagraphs(doc, content).Item.getAttribute('Numbered'), 'true');
});

test('list attributes from numbering.xml apply inside table cells', async () => {
    const { doc } = await convertBody(table([[listItem('Cell bullet', 3), p('<w:r><w:t>Other</w:t></w:r>')]]), {
        numbering: abstractNum(0, lvl(0, 'bullet', '')) + num(3, 0)
    });

    const para = doc.querySelector('elements cell paragraph');
    assert.equal(para.getAttribute('Bulleted'), 'true');
    assert.equal(para.getAttribute('Numbered'), null);
});

test('a reused parser does not carry list definitions into the next document', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const body = listItem('Item', 1);

    const first = await parser.parse(await buildDocx(body, { numbering: abstractNum(0, lvl(0, 'bullet', '')) + num(1, 0) }));
    assert.equal(first.elements[0].numbering.format, 'bullet');

    const second = await parser.parse(await buildDocx(body));
    assert.equal(second.elements[0].numbering.format, undefined);
    assert.equal(second.elements[0].numbering.numId, '1');
});

const run = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${text}</w:t></w:r>`;
const styled = (styleId, runs, pPr = '') => p(runs, `<w:pPr><w:pStyle w:val="${styleId}"/>${pPr}</w:pPr>`);
const paragraphStyle = (id, { basedOn, pPr = '', rPr = '', isDefault = false } = {}) =>
    `<w:style w:type="paragraph"${isDefault ? ' w:default="1"' : ''} w:styleId="${id}">` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : '') +
    (pPr ? `<w:pPr>${pPr}</w:pPr>` : '') + (rPr ? `<w:rPr>${rPr}</w:rPr>` : '') + '</w:style>';
const characterStyle = (id, rPr) => `<w:style w:type="character" w:styleId="${id}"><w:rPr>${rPr}</w:rPr></w:style>`;

/** The content element holding exactly this text */
function contentWithText(doc, content, text) {
    const found = contentElements(doc).filter(el => textOf(el, content) === text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} elements`);
    return found[0];
}

/** The character attributes the generator wrote for a piece of text */
function charAttrs(doc, content, text) {
    const el = contentWithText(doc, content, text);
    return Object.fromEntries(['family', 'size', 'bold', 'italic', 'underline', 'strikethrough', 'foreground']
        .map(name => [name, el.getAttribute(name)]));
}

const PLAIN = { family: 'Times New Roman', size: '12', bold: null, italic: null, underline: null, strikethrough: null, foreground: null };

test('w:docDefaults font and size apply to unformatted runs', async () => {
    const { doc, content } = await convertBody(p(run('Plain')) + p(run('Direct', '<w:rFonts w:ascii="Courier New"/>')), {
        styles: '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>' +
            '</w:rPr></w:rPrDefault></w:docDefaults>'
    });

    assert.deepEqual(charAttrs(doc, content, 'Plain'), { ...PLAIN, family: 'Arial', size: '11' });
    assert.deepEqual(charAttrs(doc, content, 'Direct'), { ...PLAIN, family: 'Courier New', size: '11' });
});

test('a w:basedOn chain is followed and a later style can turn a property off', async () => {
    const { doc, content } = await convertBody(
        styled('Child', run('Child text')) + styled('Grandchild', run('Grandchild text')), {
            styles: paragraphStyle('Base', { rPr: '<w:b/><w:sz w:val="28"/>' }) +
                paragraphStyle('Child', { basedOn: 'Base', rPr: '<w:i/>' }) +
                paragraphStyle('Grandchild', { basedOn: 'Child', rPr: '<w:b w:val="0"/><w:u w:val="single"/>' })
        });

    assert.deepEqual(charAttrs(doc, content, 'Child text'), { ...PLAIN, size: '14', bold: 'true', italic: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Grandchild text'), { ...PLAIN, size: '14', italic: 'true', underline: 'true' });
});

test('a style without w:type is a paragraph style', async () => {
    const { doc, content } = await convertBody(styled('Untyped', run('Untyped text')), {
        styles: '<w:style w:styleId="Untyped"><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
    });

    assert.deepEqual(charAttrs(doc, content, 'Untyped text'), { ...PLAIN, size: '14', bold: 'true' });
});

test('a cyclic, self-referencing or dangling w:basedOn does not throw or hang', async () => {
    const { doc, content } = await convertBody(
        styled('A', run('Cycle')) + styled('Self', run('Self')) + styled('Dangling', run('Dangling')), {
            styles: paragraphStyle('A', { basedOn: 'B', rPr: '<w:b/>' }) +
                paragraphStyle('B', { basedOn: 'A', rPr: '<w:i/><w:b w:val="0"/>' }) +
                paragraphStyle('Self', { basedOn: 'Self', rPr: '<w:strike/>' }) +
                paragraphStyle('Dangling', { basedOn: 'Missing', rPr: '<w:sz w:val="32"/>' })
        });

    // The chain stops where it loops: A takes B underneath, then its own bold on top
    assert.deepEqual(charAttrs(doc, content, 'Cycle'), { ...PLAIN, bold: 'true', italic: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Self'), { ...PLAIN, strikethrough: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Dangling'), { ...PLAIN, size: '16' });
});

test('a character style applies over the paragraph style, and direct w:rPr wins over both', async () => {
    const { doc, content, parsed } = await convertBody(styled('Heading',
        run('Heading') +
        run('Small', '<w:rStyle w:val="Small"/>') +
        run('Small but italic', '<w:rStyle w:val="Small"/><w:i/>') +
        run('Not bold', '<w:b w:val="0"/><w:color w:val="auto"/>') +
        run('Unknown style', '<w:rStyle w:val="Nope"/>')), {
        styles: paragraphStyle('Heading', { rPr: '<w:b/><w:sz w:val="28"/><w:color w:val="FF0000"/>' }) +
            characterStyle('Small', '<w:sz w:val="20"/><w:u w:val="single"/><w:color w:val="0000FF"/>')
    });

    const red = String(0xFFFF0000 | 0);
    const blue = String(0xFF0000FF | 0);
    assert.deepEqual(charAttrs(doc, content, 'Heading'), { ...PLAIN, size: '14', bold: 'true', foreground: red });
    assert.deepEqual(charAttrs(doc, content, 'Small'), { ...PLAIN, size: '10', bold: 'true', underline: 'true', foreground: blue });
    assert.deepEqual(charAttrs(doc, content, 'Small but italic'),
        { ...PLAIN, size: '10', bold: 'true', italic: 'true', underline: 'true', foreground: blue });
    assert.deepEqual(charAttrs(doc, content, 'Not bold'), { ...PLAIN, size: '14' });
    assert.equal('color' in parsed.elements[0].runs[3], false);
    assert.deepEqual(charAttrs(doc, content, 'Unknown style'), { ...PLAIN, size: '14', bold: 'true', foreground: red });
});

test('bold, italic and strike toggle between the paragraph and character styles', async () => {
    const styles = paragraphStyle('Heading', { rPr: '<w:b/><w:i/>' }) +
        characterStyle('Strong', '<w:b/>') +
        characterStyle('NotBold', '<w:b w:val="0"/>') +
        characterStyle('StrongBase', '<w:b/><w:strike/>') +
        '<w:style w:type="character" w:styleId="Derived"><w:basedOn w:val="StrongBase"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>';
    const { doc, content } = await convertBody(
        styled('Heading',
            run('Strong in heading', '<w:rStyle w:val="Strong"/>') +
            run('Off in style', '<w:rStyle w:val="NotBold"/>') +
            run('Direct bold', '<w:rStyle w:val="Strong"/><w:b/>') +
            run('Derived', '<w:rStyle w:val="Derived"/>')) +
        p(run('Strong in body', '<w:rStyle w:val="Strong"/>')), { styles });

    // On in both styles is off; off in a style leaves it unchanged; direct formatting is absolute
    assert.deepEqual(charAttrs(doc, content, 'Strong in heading'), { ...PLAIN, italic: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Off in style'), { ...PLAIN, bold: 'true', italic: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Direct bold'), { ...PLAIN, bold: 'true', italic: 'true' });
    // Within one style level the basedOn chain overrides: Derived turns its base's bold off
    assert.deepEqual(charAttrs(doc, content, 'Derived'), { ...PLAIN, bold: 'true', italic: 'true', strikethrough: 'true' });
    assert.deepEqual(charAttrs(doc, content, 'Strong in body'), { ...PLAIN, bold: 'true' });

    // On in the document defaults: a style level that leaves it unset counts as on, and it
    // shows on when an even number of style levels turn it off
    const withDefaults = await convertBody(
        styled('Heading', run('Both styles on', '<w:rStyle w:val="Strong"/>')) +
        styled('Regular', run('Paragraph style off')) +
        styled('Regular', run('Character style on', '<w:rStyle w:val="Strong"/>')) +
        p(run('Character style off', '<w:rStyle w:val="NotBold"/>')) +
        p(run('No styles')), {
            styles: '<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>' + styles +
                paragraphStyle('Regular', { rPr: '<w:b w:val="0"/>' })
        });
    const bold = text => charAttrs(withDefaults.doc, withDefaults.content, text).bold;
    assert.equal(bold('Both styles on'), 'true');
    assert.equal(bold('Paragraph style off'), null);
    assert.equal(bold('Character style on'), null);
    assert.equal(bold('Character style off'), null);
    assert.equal(bold('No styles'), 'true');
});

test('paragraph alignment and indents come from the paragraph style unless set directly', async () => {
    const { doc, content } = await convertBody(
        styled('Indented', run('From style')) +
        styled('Indented', run('Direct'), '<w:jc w:val="right"/><w:ind w:left="1440"/>') +
        styled('Indented', run('Hanging'), '<w:ind w:hanging="360"/>') +
        p(run('Default style')) +
        styled('Missing', run('Unknown style')), {
            styles: paragraphStyle('Normal', { isDefault: true, pPr: '<w:jc w:val="both"/>' }) +
                paragraphStyle('Indented', { basedOn: 'Normal', pPr: '<w:jc w:val="center"/><w:ind w:left="720" w:firstLine="360"/>' })
        });

    const para = text => {
        const el = contentWithText(doc, content, text).closest('paragraph');
        return ['Alignment', 'LeftIndent', 'RightIndent', 'FirstLineIndent'].map(name => el.getAttribute(name));
    };
    assert.deepEqual(para('From style'), ['1', '36.0', '0.0', '18.0']);
    assert.deepEqual(para('Direct'), ['2', '72.0', '0.0', '18.0']);
    assert.deepEqual(para('Hanging'), ['1', '36.0', '0.0', null]);
    assert.deepEqual(para('Default style'), ['3', '0.0', '0.0', null]);
    assert.deepEqual(para('Unknown style'), ['3', '0.0', '0.0', null]);
});

// Lists from paragraph styles: w:numPr in styles.xml inherits through w:basedOn,
// and the paragraph's own w:numId and w:ilvl each override the style's

const LIST_ATTRIBUTES = ['Numbered', 'Bulleted', 'NumberType', 'BulletType', 'ListLevel', 'ListId'];
const numPr = ({ numId, ilvl } = {}) => '<w:numPr>' +
    (ilvl === undefined ? '' : `<w:ilvl w:val="${ilvl}"/>`) +
    (numId === undefined ? '' : `<w:numId w:val="${numId}"/>`) + '</w:numPr>';

/** The list attributes of the paragraph holding this text */
function listAttrs(doc, content, text) {
    const el = contentWithText(doc, content, text).closest('paragraph');
    return Object.fromEntries(LIST_ATTRIBUTES.map(name => [name, el.getAttribute(name)])
        .filter(([, value]) => value !== null));
}

const STYLE_LISTS = abstractNum(0, lvl(0, 'decimal', '%1.') + lvl(1, 'lowerLetter', '%2)')) +
    abstractNum(1, lvl(0, 'bullet', '')) + num(5, 0) + num(6, 1);

test('a paragraph style\'s w:numPr numbers the paragraph, directly or through w:basedOn', async () => {
    const { doc, content, parsed } = await convertBody(
        styled('Argument', run('Direct style')) + styled('SubArgument', run('Inherited')) +
        styled('Point', run('Bullet style')), {
            styles: paragraphStyle('Argument', { pPr: numPr({ numId: 5, ilvl: 0 }) }) +
                paragraphStyle('SubArgument', { basedOn: 'Argument', pPr: numPr({ ilvl: 1 }) }) +
                paragraphStyle('Point', { pPr: numPr({ numId: 6, ilvl: 0 }) }),
            numbering: STYLE_LISTS
        });

    assert.deepEqual(listAttrs(doc, content, 'Direct style'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_NUMBER_DOT', ListLevel: '0', ListId: '5' });
    assert.deepEqual(listAttrs(doc, content, 'Inherited'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', ListLevel: '1', ListId: '5' });
    assert.deepEqual(listAttrs(doc, content, 'Bullet style'),
        { Bulleted: 'true', BulletType: 'BULLET_TYPE_ELLIPSE', ListLevel: '0', ListId: '6' });

    // Only the numbering object carries the list; no style-only keys leak into the paragraph
    const first = parsed.elements[0];
    // Copied into this realm: the parser's objects come from the vm context
    assert.deepEqual({ ...first.numbering }, { level: 0, numId: '5', format: 'decimal', lvlText: '%1.' });
    for (const key of ['numId', 'level', 'ilvl', 'numPr']) assert.equal(first[key], undefined, key);
});

test('the paragraph\'s own w:ilvl and w:numId override the style\'s, each on its own', async () => {
    const { doc, content } = await convertBody(
        styled('Argument', run('Own level'), numPr({ ilvl: 1 })) +
        styled('Argument', run('Own list'), numPr({ numId: 6 })) +
        styled('Argument', run('Own both'), numPr({ numId: 6, ilvl: 0 })), {
            styles: paragraphStyle('Argument', { pPr: numPr({ numId: 5, ilvl: 0 }) }),
            numbering: STYLE_LISTS
        });

    assert.deepEqual(listAttrs(doc, content, 'Own level'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', ListLevel: '1', ListId: '5' });
    assert.deepEqual(listAttrs(doc, content, 'Own list'),
        { Bulleted: 'true', BulletType: 'BULLET_TYPE_ELLIPSE', ListLevel: '0', ListId: '6' });
    assert.deepEqual(listAttrs(doc, content, 'Own both'),
        { Bulleted: 'true', BulletType: 'BULLET_TYPE_ELLIPSE', ListLevel: '0', ListId: '6' });
});

test('w:numId 0 on the paragraph or a nearer style switches style numbering off', async () => {
    const { doc, content } = await convertBody(
        styled('Argument', run('Off on paragraph'), numPr({ numId: 0 })) +
        styled('Unnumbered', run('Off in style')) +
        styled('Unnumbered', run('Back on'), numPr({ numId: 5 })), {
            styles: paragraphStyle('Argument', { pPr: numPr({ numId: 5, ilvl: 1 }) }) +
                paragraphStyle('Unnumbered', { basedOn: 'Argument', pPr: numPr({ numId: 0 }) }),
            numbering: STYLE_LISTS
        });

    assert.deepEqual(listAttrs(doc, content, 'Off on paragraph'), {});
    assert.deepEqual(listAttrs(doc, content, 'Off in style'), {});
    // The level still inherits from Argument underneath the switched-off style
    assert.deepEqual(listAttrs(doc, content, 'Back on'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', ListLevel: '1', ListId: '5' });
});

test('a w:numPr with w:numId but no w:ilvl, on the paragraph or a style, is level 0', async () => {
    const { doc, content } = await convertBody(
        p(run('Paragraph without ilvl'), `<w:pPr>${numPr({ numId: 5 })}</w:pPr>`) +
        styled('NoLevel', run('Style without ilvl')), {
            styles: paragraphStyle('NoLevel', { pPr: numPr({ numId: 5 }) }),
            numbering: STYLE_LISTS
        });

    for (const text of ['Paragraph without ilvl', 'Style without ilvl']) {
        assert.deepEqual(listAttrs(doc, content, text),
            { Numbered: 'true', NumberType: 'NUMBER_TYPE_NUMBER_DOT', ListLevel: '0', ListId: '5' }, text);
    }
});

test('a w:ilvl without any w:numId is not a list item', async () => {
    const { doc, content } = await convertBody(
        p(run('Level only'), `<w:pPr>${numPr({ ilvl: 1 })}</w:pPr>`) + styled('LevelOnly', run('Style level only')), {
            styles: paragraphStyle('LevelOnly', { pPr: numPr({ ilvl: 1 }) }),
            numbering: STYLE_LISTS
        });

    assert.deepEqual(listAttrs(doc, content, 'Level only'), {});
    assert.deepEqual(listAttrs(doc, content, 'Style level only'), {});
});

test('style numbering survives a looping or dangling w:basedOn and a numId missing from numbering.xml', async () => {
    const { doc, content } = await convertBody(
        styled('A', run('Cycle')) + styled('Dangling', run('Dangling')) +
        styled('Unknown', run('Odd unknown')) + styled('UnknownEven', run('Even unknown')), {
            styles: paragraphStyle('A', { basedOn: 'B', pPr: numPr({ ilvl: 1 }) }) +
                paragraphStyle('B', { basedOn: 'A', pPr: numPr({ numId: 5, ilvl: 0 }) }) +
                paragraphStyle('Dangling', { basedOn: 'Missing', pPr: numPr({ numId: 6 }) }) +
                paragraphStyle('Unknown', { pPr: numPr({ numId: 7, ilvl: 0 }) }) +
                paragraphStyle('UnknownEven', { pPr: numPr({ numId: 8, ilvl: 0 }) }),
            numbering: STYLE_LISTS
        });

    assert.deepEqual(listAttrs(doc, content, 'Cycle'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', ListLevel: '1', ListId: '5' });
    assert.deepEqual(listAttrs(doc, content, 'Dangling'),
        { Bulleted: 'true', BulletType: 'BULLET_TYPE_ELLIPSE', ListLevel: '0', ListId: '6' });
    // The same numId guess a paragraph w:numPr without a list definition gets
    assert.deepEqual(listAttrs(doc, content, 'Odd unknown'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_NUMBER_TRE', ListLevel: '0', ListId: '7' });
    assert.deepEqual(listAttrs(doc, content, 'Even unknown'),
        { Bulleted: 'true', BulletType: 'BULLET_TYPE_ELLIPSE', ListLevel: '0', ListId: '8' });
});

test('the default paragraph style\'s w:numPr numbers paragraphs without a w:pStyle', async () => {
    const { doc, content } = await convertBody(p(run('Unstyled')), {
        styles: paragraphStyle('Normal', { isDefault: true, pPr: numPr({ numId: 5, ilvl: 0 }) }),
        numbering: STYLE_LISTS
    });

    assert.deepEqual(listAttrs(doc, content, 'Unstyled'),
        { Numbered: 'true', NumberType: 'NUMBER_TYPE_NUMBER_DOT', ListLevel: '0', ListId: '5' });
});

// Paragraph spacing: w:spacing -> SpaceAbove, SpaceBelow and LineSpacing

const SPACING_ATTRIBUTES = ['SpaceAbove', 'SpaceBelow', 'LineSpacing'];

/** The spacing attributes of the paragraph holding this text */
function spacingOf(doc, content, text) {
    const el = contentWithText(doc, content, text).closest('paragraph');
    return Object.fromEntries(SPACING_ATTRIBUTES.map(name => [name, el.getAttribute(name)]));
}

const NO_SPACING = { SpaceAbove: null, SpaceBelow: null, LineSpacing: null };

test('w:before and w:after become SpaceAbove and SpaceBelow in points', async () => {
    const { doc, content } = await convertBody(
        p(run('Spaced'), '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>') +
        p(run('Half points'), '<w:pPr><w:spacing w:before="30" w:after="0"/></w:pPr>') +
        p(run('Measure'), '<w:pPr><w:spacing w:after="6pt"/></w:pPr>'));

    assert.deepEqual(spacingOf(doc, content, 'Spaced'), { ...NO_SPACING, SpaceAbove: '12.0', SpaceBelow: '6.0' });
    assert.deepEqual(spacingOf(doc, content, 'Half points'), { ...NO_SPACING, SpaceAbove: '1.5' });
    assert.deepEqual(spacingOf(doc, content, 'Measure'), { ...NO_SPACING, SpaceBelow: '6.0' });
});

test('auto line spacing becomes the fraction over single spacing', async () => {
    const { doc, content } = await convertBody(
        p(run('One and a half'), '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>') +
        p(run('No rule'), '<w:pPr><w:spacing w:line="360"/></w:pPr>') +
        p(run('Multiple'), '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr>') +
        p(run('Double'), '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>') +
        p(run('Single'), '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>') +
        p(run('Tighter'), '<w:pPr><w:spacing w:line="200" w:lineRule="auto"/></w:pPr>'));

    assert.equal(spacingOf(doc, content, 'One and a half').LineSpacing, '0.5');
    assert.equal(spacingOf(doc, content, 'No rule').LineSpacing, '0.5');
    assert.equal(spacingOf(doc, content, 'Multiple').LineSpacing, '0.15');
    assert.equal(spacingOf(doc, content, 'Double').LineSpacing, '1.0');
    assert.deepEqual(spacingOf(doc, content, 'Single'), NO_SPACING);
    assert.deepEqual(spacingOf(doc, content, 'Tighter'), NO_SPACING);
});

test('exact and at-least line spacing write no LineSpacing, even over a style\'s auto spacing', async () => {
    const { doc, content } = await convertBody(
        styled('OneHalf', run('Exact'), '<w:spacing w:line="360" w:lineRule="exact"/>') +
        styled('OneHalf', run('At least'), '<w:spacing w:line="480" w:lineRule="atLeast"/>') +
        styled('OneHalf', run('From style')), {
            styles: paragraphStyle('OneHalf', { pPr: '<w:spacing w:line="360" w:lineRule="auto"/>' })
        });

    assert.deepEqual(spacingOf(doc, content, 'Exact'), NO_SPACING);
    assert.deepEqual(spacingOf(doc, content, 'At least'), NO_SPACING);
    assert.equal(spacingOf(doc, content, 'From style').LineSpacing, '0.5');
});

test('the line height and its rule are inherited separately', async () => {
    const { doc, content } = await convertBody(
        styled('OneHalf', run('Direct exact rule'), '<w:spacing w:lineRule="exact"/>') +
        styled('Exact', run('Direct line under an exact rule'), '<w:spacing w:line="480"/>') +
        styled('Exact', run('Direct auto rule'), '<w:spacing w:lineRule="auto"/>'), {
            styles: paragraphStyle('OneHalf', { pPr: '<w:spacing w:line="360" w:lineRule="auto"/>' }) +
                paragraphStyle('Exact', { pPr: '<w:spacing w:line="360" w:lineRule="exact"/>' })
        });

    assert.deepEqual(spacingOf(doc, content, 'Direct exact rule'), NO_SPACING);
    assert.deepEqual(spacingOf(doc, content, 'Direct line under an exact rule'), NO_SPACING);
    assert.equal(spacingOf(doc, content, 'Direct auto rule').LineSpacing, '0.5');
});

test('spacing merges per attribute through docDefaults, the w:basedOn chain and direct w:pPr', async () => {
    const { doc, content } = await convertBody(
        styled('Heading', run('Style before, direct after'), '<w:spacing w:after="120"/>') +
        styled('Heading', run('Direct after zero'), '<w:spacing w:after="0"/>') +
        styled('Heading', run('Style only')) +
        p(run('Defaults only')), {
            styles: '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/>' +
                '</w:pPr></w:pPrDefault></w:docDefaults>' +
                paragraphStyle('Base', { pPr: '<w:spacing w:before="240"/>' }) +
                paragraphStyle('Heading', { basedOn: 'Base', pPr: '<w:spacing w:after="240"/>' })
        });

    assert.deepEqual(spacingOf(doc, content, 'Style before, direct after'),
        { SpaceAbove: '12.0', SpaceBelow: '6.0', LineSpacing: '0.08' });
    assert.deepEqual(spacingOf(doc, content, 'Direct after zero'),
        { SpaceAbove: '12.0', SpaceBelow: null, LineSpacing: '0.08' });
    assert.deepEqual(spacingOf(doc, content, 'Style only'),
        { SpaceAbove: '12.0', SpaceBelow: '12.0', LineSpacing: '0.08' });
    assert.deepEqual(spacingOf(doc, content, 'Defaults only'),
        { SpaceAbove: null, SpaceBelow: '8.0', LineSpacing: '0.08' });
});

test('an unreadable or negative spacing value leaves the inherited value in place', async () => {
    const { doc, content } = await convertBody(
        styled('Spaced', run('Bad values'), '<w:spacing w:before="abc" w:after="-120" w:line="x"/>'), {
            styles: paragraphStyle('Spaced', { pPr: '<w:spacing w:before="240" w:after="120" w:line="360"/>' })
        });

    assert.deepEqual(spacingOf(doc, content, 'Bad values'), { SpaceAbove: '12.0', SpaceBelow: '6.0', LineSpacing: '0.5' });
});

test('paragraphs inside table cells get spacing too', async () => {
    const { doc, content } = await convertBody(table([[
        p(run('Cell spaced'), '<w:pPr><w:spacing w:before="240" w:after="120" w:line="360"/></w:pPr>'),
        p(run('Cell plain'))
    ]]));

    assert.deepEqual(spacingOf(doc, content, 'Cell spaced'), { SpaceAbove: '12.0', SpaceBelow: '6.0', LineSpacing: '0.5' });
    assert.deepEqual(spacingOf(doc, content, 'Cell plain'), NO_SPACING);
});

test('the part of a paragraph after a page break does not repeat the space above', async () => {
    const { doc, content } = await convertBody(p(
        '<w:r><w:t>X</w:t><w:br w:type="page"/><w:t>Y</w:t></w:r>',
        '<w:pPr><w:spacing w:before="240" w:after="120" w:line="360"/></w:pPr>'));

    assert.deepEqual(spacingOf(doc, content, 'X'), { SpaceAbove: '12.0', SpaceBelow: '6.0', LineSpacing: '0.5' });
    assert.deepEqual(spacingOf(doc, content, 'Y'), { SpaceAbove: null, SpaceBelow: '6.0', LineSpacing: '0.5' });
});

test('a document without w:spacing writes no spacing attributes', async () => {
    const { xml } = await convertBody(
        p(run('Plain')) + p(run('Indented'), '<w:pPr><w:jc w:val="both"/><w:ind w:left="720"/></w:pPr>') +
        table([[p(run('Cell')), '<w:p/>']]), {
            styles: paragraphStyle('Normal', { isDefault: true, pPr: '<w:jc w:val="center"/>' })
        });

    for (const name of SPACING_ATTRIBUTES) assert.ok(!xml.includes(name), `unexpected ${name}`);
});

test('without styles.xml, or with a malformed one, runs keep the built-in defaults', async (t) => {
    const body = styled('Heading', run('Text')) + p(run('Bold', '<w:b/>'), '<w:pPr><w:jc w:val="center"/></w:pPr>');
    const warn = t.mock.method(console, 'warn', () => {});

    for (const options of [{}, { styles: '' }, { styles: '<w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:b/>' }]) {
        const { doc, content } = await convertBody(body, options);
        assert.deepEqual(charAttrs(doc, content, 'Text'), PLAIN);
        assert.deepEqual(charAttrs(doc, content, 'Bold'), { ...PLAIN, bold: 'true' });
        assert.equal(contentWithText(doc, content, 'Text').closest('paragraph').getAttribute('Alignment'), '0');
        assert.equal(contentWithText(doc, content, 'Bold').closest('paragraph').getAttribute('Alignment'), '1');
    }
    // Only the malformed part warns
    assert.equal(warn.mock.calls.filter(call => String(call.arguments[0]).includes('styles')).length, 1);
});

test('a reused parser does not carry styles into the next document', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const body = styled('Heading', run('Text'));

    const first = await parser.parse(await buildDocx(body, { styles: paragraphStyle('Heading', { rPr: '<w:b/>' }) }));
    assert.equal(first.elements[0].runs[0].bold, true);

    const second = await parser.parse(await buildDocx(body));
    assert.equal(second.elements[0].runs[0].bold, false);
});

// Page setup: the body's final w:sectPr -> <pageFormat>

const TODAY_PAGE_FORMAT = '<pageFormat mediaSizeName="1" leftMargin="42.51968479156494" rightMargin="28.34645652770996" ' +
    'topMargin="14.17322826385498" bottomMargin="14.17322826385498" paperOrientation="1" headerFOffset="20.0" footerFOffset="20.0" />';
const PAGE_ATTRIBUTES = ['leftMargin', 'rightMargin', 'topMargin', 'bottomMargin', 'paperOrientation', 'headerFOffset', 'footerFOffset'];
const ONE_INCH = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';

async function pageFormat(sectPr, options = {}) {
    const { doc, xml } = await convertBody(p('<w:r><w:t>Text</w:t></w:r>'), { ...options, sectPr });
    const page = doc.querySelector('properties > pageFormat');
    return { page, xml, attr: name => Number(page.getAttribute(name)) };
}

const assertClose = (actual, expected, message) =>
    assert.ok(Math.abs(actual - expected) < 1e-9, `${message}: ${actual}, expected ${expected}`);

test('a document without w:sectPr keeps the previous page format exactly', async () => {
    for (const sectPr of [undefined, '']) {
        const { xml } = await pageFormat(sectPr);
        assert.ok(xml.includes(`<properties>${TODAY_PAGE_FORMAT}</properties>`), xml.split('\n')[3]);
    }
});

test('w:pgMar margins and header/footer offsets are converted from twips to points', async () => {
    const { page, attr } = await pageFormat(`<w:pgSz w:w="11906" w:h="16838"/>${ONE_INCH}`);
    for (const side of ['top', 'right', 'bottom', 'left']) assert.equal(attr(`${side}Margin`), 72, side);
    assert.equal(attr('headerFOffset'), 36);
    assert.equal(attr('footerFOffset'), 36);
    assert.equal(page.getAttribute('paperOrientation'), '1');
    assert.equal(page.getAttribute('mediaSizeName'), '1');

    const uneven = await pageFormat('<w:pgMar w:top="1417" w:right="1134" w:bottom="851" w:left="1701" w:header="709" w:footer="567"/>');
    assertClose(uneven.attr('topMargin'), 70.85, 'top');
    assertClose(uneven.attr('rightMargin'), 56.7, 'right');
    assertClose(uneven.attr('bottomMargin'), 42.55, 'bottom');
    assertClose(uneven.attr('leftMargin'), 85.05, 'left');
    assertClose(uneven.attr('headerFOffset'), 35.45, 'header');
    assertClose(uneven.attr('footerFOffset'), 28.35, 'footer');
});

test('the gutter widens the left margin, or the top or right one when Word puts it there', async () => {
    const pgMar = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="567"/>';

    const left = await pageFormat(pgMar);
    assertClose(left.attr('leftMargin'), 72 + 28.35, 'left');
    assert.equal(left.attr('rightMargin'), 72);
    assert.equal(left.attr('topMargin'), 72);

    const top = await pageFormat(pgMar, { settings: '<w:gutterAtTop/>' });
    assertClose(top.attr('topMargin'), 72 + 28.35, 'top');
    assert.equal(top.attr('leftMargin'), 72);

    const right = await pageFormat(`${pgMar}<w:rtlGutter/>`);
    assertClose(right.attr('rightMargin'), 72 + 28.35, 'right');
    assert.equal(right.attr('leftMargin'), 72);

    const off = await pageFormat(pgMar, { settings: '<w:gutterAtTop w:val="false"/>' });
    assertClose(off.attr('leftMargin'), 72 + 28.35, 'left with gutterAtTop off');
});

test('a negative top or bottom margin uses its absolute value', async () => {
    const { attr } = await pageFormat('<w:pgMar w:top="-1440" w:right="1440" w:bottom="-720" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>');
    assert.equal(attr('topMargin'), 72);
    assert.equal(attr('bottomMargin'), 36);
});

test('landscape comes from w:orient or from a page wider than it is tall', async () => {
    const orient = await pageFormat(`<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>${ONE_INCH}`);
    assert.equal(orient.page.getAttribute('paperOrientation'), '0');

    const orientOnly = await pageFormat('<w:pgSz w:orient="landscape"/>');
    assert.equal(orientOnly.page.getAttribute('paperOrientation'), '0');

    const wide = await pageFormat(`<w:pgSz w:w="16838" w:h="11906"/>${ONE_INCH}`);
    assert.equal(wide.page.getAttribute('paperOrientation'), '0');

    const tall = await pageFormat(`<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>${ONE_INCH}`);
    assert.equal(tall.page.getAttribute('paperOrientation'), '1');
});

test('only the body\'s final w:sectPr is read, not a paragraph\'s section break', async () => {
    const sectionBreak = p('<w:r><w:t>First section</w:t></w:r>',
        '<w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
        '<w:pgMar w:top="2880" w:right="2880" w:bottom="2880" w:left="2880" w:header="0" w:footer="0"/></w:sectPr></w:pPr>');
    const { doc } = await convertBody(sectionBreak + p('<w:r><w:t>Last section</w:t></w:r>'),
        { sectPr: `<w:pgSz w:w="11906" w:h="16838"/>${ONE_INCH}` });
    const page = doc.querySelector('properties > pageFormat');
    assert.equal(page.getAttribute('paperOrientation'), '1');
    assert.equal(Number(page.getAttribute('topMargin')), 72);
});

test('universal measures such as "2.5cm" are converted to points', async () => {
    const { attr } = await pageFormat('<w:pgMar w:top="2.5cm" w:right="1in" w:bottom="25mm" w:left="72pt" w:header="0.5in" w:footer="3pc"/>');
    assertClose(attr('topMargin'), 2.5 * 72 / 2.54, 'top');
    assert.equal(attr('rightMargin'), 72);
    assertClose(attr('bottomMargin'), 25 * 72 / 25.4, 'bottom');
    assert.equal(attr('leftMargin'), 72);
    assert.equal(attr('headerFOffset'), 36);
    assert.equal(attr('footerFOffset'), 36);
});

test('a missing or malformed page setup value falls back to the default for that attribute only', async () => {
    const fallback = name => new RegExp(`${name}="([^"]*)"`).exec(TODAY_PAGE_FORMAT)[1];

    const { page, xml } = await pageFormat(
        '<w:pgSz w:w="wide" w:h=""/>' +
        '<w:pgMar w:top="abc" w:right="1440" w:bottom="" w:left="-1440" w:header="12px" w:gutter="NaN"/>');
    assert.equal(page.getAttribute('topMargin'), fallback('topMargin'));
    assert.equal(page.getAttribute('bottomMargin'), fallback('bottomMargin'));
    assert.equal(page.getAttribute('leftMargin'), fallback('leftMargin'));
    assert.equal(page.getAttribute('headerFOffset'), fallback('headerFOffset'));
    assert.equal(page.getAttribute('footerFOffset'), fallback('footerFOffset'));
    assert.equal(page.getAttribute('paperOrientation'), '1');
    assert.equal(Number(page.getAttribute('rightMargin')), 72);

    for (const name of PAGE_ATTRIBUTES) {
        const value = page.getAttribute(name);
        assert.ok(value && Number.isFinite(Number(value)), `${name}="${value}"`);
    }
    assert.ok(!xml.includes('NaN'));
    assert.ok(!/="\s*"/.test(xml.split('\n')[3]), 'no empty attribute in <pageFormat>');
});

test('a gutter is not added to a margin that fell back to its default', async () => {
    const { page } = await pageFormat('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:gutter="567"/>');
    assert.equal(page.getAttribute('leftMargin'), '42.51968479156494');
});

test('a reused parser does not carry page setup into the next document', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const body = p('<w:r><w:t>Text</w:t></w:r>');

    const first = await parser.parse(await buildDocx(body, { sectPr: `<w:pgSz w:w="16838" w:h="11906"/>${ONE_INCH}` }));
    assert.equal(first.pageSetup.landscape, true);
    assert.equal(first.pageSetup.topMargin, 72);

    const second = await parser.parse(await buildDocx(body));
    assert.deepEqual({ ...second.pageSetup }, {});
});
