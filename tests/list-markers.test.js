const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries, loadConverter } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

// List markers. A bullet level is written as the documented BulletType that draws its bullet
// character (Docs.md, "Madde İşareti Türleri"), read through the level's symbol font as w:sym is;
// any other bullet stays BULLET_TYPE_ELLIPSE. Numbered levels keep today's NumberType. A marker the
// UDF can't keep (another bullet, a picture bullet, number text around or beyond the level's own
// number, another number format, a start other than 1) is reported as a listMarkers conversion note.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const numPr = (numId, ilvl = 0) => `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
const item = (text, numId, ilvl = 0) => p(r(text), numPr(numId, ilvl));

/** A w:lvl; font goes in its w:rPr/w:rFonts, as Word writes a bullet level's font */
const lvl = (ilvl, numFmt, lvlText, { start = 1, font = '', extra = '' } = {}) =>
    `<w:lvl w:ilvl="${ilvl}">${start === null ? '' : `<w:start w:val="${start}"/>`}<w:numFmt w:val="${numFmt}"/>` +
    `<w:lvlText w:val="${lvlText}"/>${extra}<w:lvlJc w:val="left"/>` +
    (font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : '') + '</w:lvl>';
const abstractNum = (id, levels) => `<w:abstractNum w:abstractNumId="${id}">${levels}</w:abstractNum>`;
const num = (numId, abstractNumId, overrides = '') =>
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/>${overrides}</w:num>`;
/** One list per level definition: list i + 1 uses abstractNum i, whose level ilvl is the given one */
const lists = levels => levels.map(([ilvl, ...args], i) => abstractNum(i, lvl(ilvl, ...args))).join('') +
    levels.map((_, i) => num(i + 1, i)).join('');

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;

/** The <paragraph> holding this text */
const paragraphOf = (doc, content, text) => {
    const el = [...doc.querySelectorAll('content')].find(c => textOf(c, content) === text);
    assert.ok(el, `content ${text}`);
    return el.closest('paragraph');
};

const DOCUMENTED_BULLET_TYPES = ['BULLET_TYPE_ELLIPSE', 'BULLET_TYPE_RECTANGLE', 'BULLET_TYPE_ARROW', 'BULLET_TYPE_DIAMOND',
    'BULLET_TYPE_DIAMOND_2', 'BULLET_TYPE_TRIANGLE', 'BULLET_TYPE_RECTANGLE_D'];

test('the issue\'s reproduction: bullets map to the documented BulletType, and markers the UDF can\'t keep are reported', async () => {
    const cases = [
        ['Kare', [0, 'bullet', '', { font: 'Wingdings' }], 'BulletType', 'BULLET_TYPE_RECTANGLE', false],
        ['Tire', [0, 'bullet', '-'], 'BulletType', 'BULLET_TYPE_ELLIPSE', true],
        ['Madde', [0, 'decimal', 'MADDE %1:'], 'NumberType', 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Parantez', [0, 'decimal', '(%1)'], 'NumberType', 'NUMBER_TYPE_NUMBER_PARENTHESIS', true],
        ['Dörtten', [0, 'decimal', '%1.', { start: 4 }], 'NumberType', 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Sıra', [0, 'ordinal', '%1.'], 'NumberType', 'NUMBER_TYPE_NUMBER_DOT', true]
    ];
    const { doc, content, parsed } = await convertBody(cases.map(([text], i) => item(text, i + 1)).join(''),
        { numbering: lists(cases.map(([, level]) => level)) });

    for (const [text, , attribute, value] of cases) {
        assert.equal(paragraphOf(doc, content, text).getAttribute(attribute), value, text);
    }
    const reported = cases.filter(([, , , , note]) => note).map(([text]) => text);
    assert.deepEqual(findingSummaries(parsed.findings),
        [{ code: 'listMarkers', count: reported.length, examples: reported.slice(0, 3) }]);
    assert.deepEqual(occurrences(parsed, 'listMarkers')[0], { example: 'Tire', location: { kind: 'paragraph', paragraph: 2 } });
});

test('bullet characters, read through their symbol font, map to the documented BulletType without a note', async () => {
    const cases = [
        ['Wingdings F0A7', ['', 'Wingdings'], 'BULLET_TYPE_RECTANGLE'],
        ['Wingdings 00A7', ['§', 'Wingdings'], 'BULLET_TYPE_RECTANGLE'],
        ['Symbol F0B7', ['', 'Symbol'], 'BULLET_TYPE_ELLIPSE'],
        ['Symbol 00B7', ['·', 'Symbol'], 'BULLET_TYPE_ELLIPSE'],
        ['plain bullet', ['•', ''], 'BULLET_TYPE_ELLIPSE'],
        ['black circle', ['●', 'Arial'], 'BULLET_TYPE_ELLIPSE'],
        ['black square', ['■', ''], 'BULLET_TYPE_RECTANGLE'],
        ['diamond', ['◆', ''], 'BULLET_TYPE_DIAMOND'],
        ['Wingdings v', ['v', 'Wingdings'], 'BULLET_TYPE_DIAMOND'],
        ['lozenge', ['◊', ''], 'BULLET_TYPE_DIAMOND_2'],
        ['triangle', ['▲', ''], 'BULLET_TYPE_TRIANGLE'],
        ['white square', ['□', ''], 'BULLET_TYPE_RECTANGLE_D'],
        ['Wingdings q', ['q', 'Wingdings'], 'BULLET_TYPE_RECTANGLE_D'],
        ['Wingdings F0D8', ['', 'Wingdings'], 'BULLET_TYPE_ARROW'],
        ['arrowhead', ['➤', ''], 'BULLET_TYPE_ARROW'],
        ['pointer', ['►', ''], 'BULLET_TYPE_ARROW']
    ];
    const { doc, content, parsed } = await convertBody(cases.map(([text], i) => item(text, i + 1)).join(''),
        { numbering: lists(cases.map(([, [lvlText, font]]) => [0, 'bullet', lvlText, { font }])) });

    for (const [text, , type] of cases) {
        const para = paragraphOf(doc, content, text);
        assert.equal(para.getAttribute('Bulleted'), 'true', text);
        assert.equal(para.getAttribute('BulletType'), type, text);
        assert.equal(para.getAttribute('Numbered'), null, text);
    }
    assert.equal(finding(parsed, 'listMarkers'), undefined);
});

test('a level\'s font is read from w:ascii, else w:hAnsi', async () => {
    const level = font => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>` +
        `<w:rPr><w:rFonts ${font}/></w:rPr></w:lvl>`;
    const { doc, content, parsed } = await convertBody(item('hAnsi', 1) + item('ascii', 2), {
        numbering: abstractNum(0, level('w:hAnsi="Wingdings"')) + abstractNum(1, level('w:ascii="Wingdings" w:hAnsi="Symbol"')) +
            num(1, 0) + num(2, 1)
    });
    assert.equal(paragraphOf(doc, content, 'hAnsi').getAttribute('BulletType'), 'BULLET_TYPE_RECTANGLE');
    assert.equal(paragraphOf(doc, content, 'ascii').getAttribute('BulletType'), 'BULLET_TYPE_RECTANGLE');
    assert.equal(finding(parsed, 'listMarkers'), undefined);
});

test('a bullet with no documented type, and a picture bullet, are written as ELLIPSE and reported once per paragraph', async () => {
    const cases = [
        ['Tire', ['-', '']],
        ['Courier o', ['o', 'Courier New']],
        ['Resim', ['', 'Symbol', '<w:lvlPicBulletId w:val="0"/>']],
        ['Onay', ['', 'Wingdings']],
        ['Boş', ['', '']],
        ['Bilinmeyen', ['', 'Wingdings']],
        ['İki', ['••', '']]
    ];
    // Two paragraphs in each list: each is counted
    const { doc, content, parsed } = await convertBody(
        cases.map(([text], i) => item(text, i + 1) + item(`${text} 2`, i + 1)).join(''),
        { numbering: lists(cases.map(([, [lvlText, font, extra = '']]) => [0, 'bullet', lvlText, { font, extra }])) });

    for (const [text] of cases) {
        for (const each of [text, `${text} 2`]) {
            assert.equal(paragraphOf(doc, content, each).getAttribute('BulletType'), 'BULLET_TYPE_ELLIPSE', each);
        }
    }
    const found = finding(parsed, 'listMarkers');
    assert.equal(found.count, cases.length * 2);
    assert.deepEqual(found.examples, ['Tire', 'Tire 2', 'Courier o']);
});

test('numbered levels keep their NumberType; marker text, multi-level numbers and other formats are reported', async () => {
    const cases = [
        // [text, level, NumberType, reported]
        ['Madde', [0, 'decimal', 'MADDE %1:'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Parantez', [0, 'decimal', '(%1)'], 'NUMBER_TYPE_NUMBER_PARENTHESIS', true],
        ['Çok düzey', [1, 'decimal', '%1.%2.'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Sıra', [0, 'ordinal', '%1.'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Sıfırlı', [0, 'decimalZero', '%1.'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Yalın', [0, 'decimal', '%1'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Paragraf işareti', [0, 'decimal', '§ %1'], 'NUMBER_TYPE_NUMBER_DOT', true],
        ['Harf tire', [0, 'lowerLetter', '%1-'], 'NUMBER_TYPE_CHAR_SMALL_DOT', true],
        ['Nokta', [0, 'decimal', '%1.'], 'NUMBER_TYPE_NUMBER_DOT', false],
        ['Kapanış', [0, 'decimal', '%1)'], 'NUMBER_TYPE_NUMBER_PARENTHESIS', false],
        ['Tire', [0, 'decimal', '%1-'], 'NUMBER_TYPE_NUMBER_TRE', false],
        ['Harf', [0, 'lowerLetter', '%1)'], 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', false],
        ['Romen', [0, 'upperRoman', '%1.'], 'NUMBER_TYPE_ROMAN_BIG_DOT', false],
        ['İkinci düzey', [1, 'lowerLetter', '%2)'], 'NUMBER_TYPE_CHAR_SMALL_PARENTHESIS', false]
    ];
    const { doc, content, parsed } = await convertBody(
        cases.map(([text, [ilvl]], i) => item(text, i + 1, ilvl)).join(''),
        { numbering: lists(cases.map(([, level]) => level)) });

    for (const [text, [ilvl], type] of cases) {
        const para = paragraphOf(doc, content, text);
        assert.equal(para.getAttribute('Numbered'), 'true', text);
        assert.equal(para.getAttribute('NumberType'), type, text);
        assert.equal(para.getAttribute('ListLevel'), String(ilvl), text);
    }
    assert.deepEqual(occurrences(parsed, 'listMarkers').map(o => o.example),
        cases.filter(([, , , reported]) => reported).map(([text]) => text).slice(0, 5));
    assert.equal(finding(parsed, 'listMarkers').count, cases.filter(([, , , reported]) => reported).length);
});

test('a start value other than 1 is reported: w:start, or the paragraph\'s list\'s w:startOverride for that level', async () => {
    const { doc, content, parsed } = await convertBody(
        item('Dörtten', 1) + item('Yeniden bir', 2) + item('Üçten', 3) + item('Düz', 4) + item('Sıfırdan', 5) + item('Belirsiz', 6), {
            numbering: abstractNum(0, lvl(0, 'decimal', '%1.', { start: 4 })) +
                abstractNum(1, lvl(0, 'decimal', '%1.')) + abstractNum(2, lvl(0, 'decimal', '%1.', { start: 0 })) +
                abstractNum(3, lvl(0, 'decimal', '%1.', { start: null })) +
                // Word's Restart at 1 on a list whose level starts at 4
                num(1, 0) + num(2, 0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>') +
                // Set Numbering Value to 3 on a level that starts at 1
                num(3, 1, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>') +
                num(4, 1) + num(5, 2) + num(6, 3)
        });

    for (const text of ['Dörtten', 'Yeniden bir', 'Üçten', 'Düz', 'Sıfırdan', 'Belirsiz']) {
        assert.equal(paragraphOf(doc, content, text).getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_DOT', text);
    }
    // A level without w:start isn't reported: Word always writes it, and nothing says it isn't 1
    assert.deepEqual(occurrences(parsed, 'listMarkers').map(o => o.example), ['Dörtten', 'Üçten', 'Sıfırdan']);
});

test('a start override on a level the list doesn\'t define makes no definition, and a level override\'s own w:start counts', async () => {
    const { doc, content, parsed } = await convertBody(item('Tanımsız', 1, 1) + item('Geçersiz kılınan', 2), {
        numbering: abstractNum(0, lvl(0, 'decimal', '%1.')) +
            num(1, 0, '<w:lvlOverride w:ilvl="1"><w:startOverride w:val="5"/></w:lvlOverride>') +
            num(2, 0, `<w:lvlOverride w:ilvl="0">${lvl(0, 'decimal', '%1.', { start: 7 })}</w:lvlOverride>`)
    });
    // Level 1 of list 1 has no definition: the numId guess, as before, and no note
    assert.equal(paragraphOf(doc, content, 'Tanımsız').getAttribute('NumberType'), 'NUMBER_TYPE_NUMBER_TRE');
    assert.deepEqual(occurrences(parsed, 'listMarkers').map(o => o.example), ['Geçersiz kılınan']);
});

test('a list missing from numbering.xml and a level without a number (numFmt none) give no note', async () => {
    for (const options of [{}, { numbering: abstractNum(0, lvl(0, 'none', '')) + num(1, 0) }]) {
        const { parsed } = await convertBody(item('Bir', 1) + item('İki', 2), options);
        assert.equal(finding(parsed, 'listMarkers'), undefined, JSON.stringify(options));
    }
});

test('list markers are reported and located in a table cell, a text box, and the header and footer', async () => {
    const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
        `<w:tr>${cells.map(inner => `<w:tc>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
    const textBox = inner => '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    const parts = [
        { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: item('Üst madde', 1) },
        { name: 'footer1.xml', root: 'ftr', type: 'footer', id: 'rIdFooter1', inner: item('Alt madde', 1) }
    ];
    const sectPr = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';
    const { doc, content, parsed } = await convertBody(
        tbl([p(r('A')), item('Hücre maddesi', 1)]) + p(textBox(item('Kutu maddesi', 1))) + item('Gövde maddesi', 2),
        { numbering: abstractNum(0, lvl(0, 'bullet', '-')) + abstractNum(1, lvl(0, 'bullet', '', { font: 'Wingdings' })) +
            num(1, 0) + num(2, 1), parts, sectPr });

    assert.equal(paragraphOf(doc, content, 'Hücre maddesi').getAttribute('BulletType'), 'BULLET_TYPE_ELLIPSE');
    assert.equal(paragraphOf(doc, content, 'Gövde maddesi').getAttribute('BulletType'), 'BULLET_TYPE_RECTANGLE');
    const found = occurrences(parsed, 'listMarkers');
    assert.deepEqual(found.map(o => o.example).sort(), ['Alt madde', 'Hücre maddesi', 'Kutu maddesi', 'Üst madde'].sort());
    assert.deepEqual(found.find(o => o.example === 'Hücre maddesi').location, { kind: 'cell', table: 1, row: 1, cell: 2 });
    assert.equal(finding(parsed, 'listMarkers').count, 4);
});

test('a body list paragraph holding only a page break isn\'t written as a paragraph, so it isn\'t reported', async () => {
    const { parsed } = await convertBody(p('<w:r><w:br w:type="page"/></w:r>', numPr(1)) + p(r('Sonra')), {
        numbering: abstractNum(0, lvl(0, 'bullet', '-')) + num(1, 0)
    });
    assert.equal(finding(parsed, 'listMarkers'), undefined);
});

test('only documented BulletType values are written', async () => {
    const characters = ['•', '■', '□', '◆', '◊', '▲', '➤', '-', 'o', '✓', ''];
    const { doc } = await convertBody(characters.map((c, i) => item(`Madde ${i}`, i + 1)).join(''),
        { numbering: lists(characters.map(c => [0, 'bullet', c])) });
    const written = [...doc.querySelectorAll('paragraph[BulletType]')].map(el => el.getAttribute('BulletType'));
    assert.equal(written.length, characters.length);
    for (const type of written) assert.ok(DOCUMENTED_BULLET_TYPES.includes(type), type);
    assert.deepEqual(new Set(written), new Set(DOCUMENTED_BULLET_TYPES));
});

test('the petition fixtures give no listMarkers note', async () => {
    for (const petition of petitions) {
        const { DocxParser } = loadConverter();
        const parsed = await new DocxParser().parse(await petition.build());
        assert.ok(!plain(parsed.findings).some(f => f.code === 'listMarkers'), petition.name);
    }
});
