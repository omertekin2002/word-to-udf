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

    // On in the document defaults stays on, whatever the styles say
    const withDefaults = await convertBody(styled('Heading', run('Defaults bold', '<w:rStyle w:val="Strong"/>')),
        { styles: '<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>' + styles });
    assert.equal(charAttrs(withDefaults.doc, withDefaults.content, 'Defaults bold').bold, 'true');
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
