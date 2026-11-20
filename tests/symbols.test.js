const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { convertBody, textOf } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
/** A w:sym element, without its run */
const symEl = (font, char) => `<w:sym${font === null ? '' : ` w:font="${font}"`} w:char="${char}"/>`;
/** A run holding one w:sym */
const sym = (font, char, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${symEl(font, char)}</w:r>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const part = (kind, inner) => ({
    name: `${kind}1.xml`,
    root: kind === 'header' ? 'hdr' : 'ftr',
    type: kind,
    id: `rId${kind === 'header' ? 'Header' : 'Footer'}1`,
    inner
});
const references = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const findingsOf = parsed => plain(parsed.findings);

/** Text of each body paragraph (header, footer and table paragraphs excluded), from the CDATA pool */
const bodyParagraphs = (doc, content) => [...doc.querySelectorAll('elements > paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));

/** The content element holding exactly this text */
const elementWithText = (doc, content, text) => {
    const found = [...doc.querySelectorAll('elements content[startOffset]')].filter(el => textOf(el, content) === text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} elements`);
    return found[0];
};

const occurrences = (text, needle) => text.split(needle).length - 1;

test('Wingdings check boxes become ☑ and □ in place, with their run\'s formatting', async () => {
    const { doc, content, parsed } = await convertBody(p(
        r('Kabul ') + sym('Wingdings', 'F0FE', '<w:b/>') + r(' Ret ') + sym('Wingdings', 'F0A8') + r(' sağ')));

    assert.deepEqual(bodyParagraphs(doc, content), ['Kabul ☑ Ret □ sağ']);
    assert.equal(content.indexOf('☑'), 'Kabul '.length);
    assert.equal(content.indexOf('□'), 'Kabul ☑ Ret '.length);
    assert.equal(elementWithText(doc, content, '☑').getAttribute('bold'), 'true');
    assert.equal(elementWithText(doc, content, '□').getAttribute('bold'), null);
    assert.deepEqual(findingsOf(parsed), [
        { code: 'symbols', count: 2, examples: ['Wingdings F0FE → ☑', 'Wingdings F0A8 → □'] }
    ]);
});

test('a symbol between texts of one run joins that run\'s text, once', async () => {
    const run = `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Tebligat: </w:t>${symEl('Wingdings', 'F078')}` +
        '<w:t xml:space="preserve"> e-tebligat</w:t></w:r>';
    const { doc, content, parsed } = await convertBody(p(run));

    assert.deepEqual(plain(parsed.elements[0].runs.map(item => [item.type, item.text, item.italic])),
        [['text', 'Tebligat: ☒ e-tebligat', true]]);
    assert.equal(elementWithText(doc, content, 'Tebligat: ☒ e-tebligat').getAttribute('italic'), 'true');
    assert.equal(occurrences(content, '☒'), 1);
});

test('the plain 00xx form of w:char maps like the F0xx form, in any letter case', async () => {
    const { doc, content } = await convertBody(
        p(sym('Wingdings', 'F0FE') + sym('Wingdings', '00FE') + sym('Wingdings', 'fe') + sym('Wingdings', 'f0a8') + sym('Wingdings', '00A8')));
    assert.deepEqual(bodyParagraphs(doc, content), ['☑☑☑□□']);
});

test('font names match ignoring case and surrounding spaces', async () => {
    const { doc, content, parsed } = await convertBody(
        p(sym('WINGDINGS', 'F0FC') + sym('wingdings 2', 'F054') + sym('Wingdings 2', 'F052') + sym(' symbol ', 'F061')));
    assert.deepEqual(bodyParagraphs(doc, content), ['✓☒☑α']);
    assert.deepEqual(findingsOf(parsed)[0].examples, ['WINGDINGS F0FC → ✓', 'wingdings 2 F054 → ☒', 'Wingdings 2 F052 → ☑']);
});

test('Symbol-font Greek letters, bullets and signs map to their symbol.txt code points', async () => {
    // [w:char, the Unicode value Adobe's symbol.txt gives that code]
    const cases = [
        ['F061', 0x03B1], ['F062', 0x03B2], ['F067', 0x03B3], ['F064', 0x03B4], ['F06D', 0x03BC], ['F070', 0x03C0],
        ['F073', 0x03C3], ['F077', 0x03C9], ['F044', 0x0394], ['F053', 0x03A3], ['F057', 0x03A9], ['F050', 0x03A0],
        ['F0B7', 0x2022], ['F0A3', 0x2264], ['F0B3', 0x2265], ['F0B1', 0x00B1], ['F0B4', 0x00D7], ['F0B9', 0x2260],
        ['F0BB', 0x2248], ['F0B0', 0x00B0], ['F0AE', 0x2192], ['F0DE', 0x21D2], ['F0A0', 0x20AC], ['F02D', 0x2212],
        ['0041', 0x0391], ['F0E1', 0x2329], ['F0F1', 0x232A]
    ];
    const { doc, content } = await convertBody(p(cases.map(([char]) => sym('Symbol', char)).join('')));
    assert.deepEqual(bodyParagraphs(doc, content), [String.fromCodePoint(...cases.map(([, code]) => code))]);
});

test('every Wingdings table entry cites its glyph in L2/11-052R, and every listed UCS value matches', () => {
    const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'docx-parser.js'), 'utf8');
    const tables = source.slice(source.indexOf("    'wingdings': {"), source.indexOf('\n};', source.indexOf("    'wingdings': {")));
    let font = null;
    let entries = 0;
    for (const line of tables.split('\n')) {
        const opening = line.match(/^ {4}'(wingdings(?: 2)?)': \{/);
        if (opening) {
            font = opening[1] === 'wingdings' ? '1' : '2';
            continue;
        }
        const entry = line.match(/^ {8}0x([0-9A-F]{2}): '(.)',? \/\/ w-(\d)(\d{3}) (.+)$/u);
        if (!entry) {
            assert.match(line, /^ {4}\},?$|^\s*$/, `unexpected line in the Wingdings tables: ${line}`);
            continue;
        }
        entries++;
        const [, code, text, set, glyph, comment] = entry;
        assert.equal(set, font, `${line}: the glyph ID is from the other font`);
        assert.equal(Number(glyph), parseInt(code, 16), `${line}: the glyph ID is not this code`);
        const cited = comment.match(/U\+([0-9A-F]{4,5})/);
        assert.ok(cited, `${line}: no Unicode value cited`);
        assert.equal(text.codePointAt(0), parseInt(cited[1], 16), `${line}: the character is not the cited one`);
    }
    assert.ok(entries >= 30, `only ${entries} entries were checked`);
});

test('an unmapped symbol writes nothing but is listed, and every symbol is counted', async () => {
    const { doc, content, parsed } = await convertBody(p(
        r('A') + sym('Webdings', 'F021') + r('B') + sym('Wingdings', 'F041') + sym('Symbol', 'F0E6') +
        sym('Wingdings', '01FE') + sym('Wingdings', 'XYZ') + sym('Wingdings', 'F0FE') + r('C')));

    assert.deepEqual(bodyParagraphs(doc, content), ['AB☑C']);
    const [finding] = findingsOf(parsed);
    assert.equal(finding.code, 'symbols');
    assert.equal(finding.count, 6);
    // Up to three examples: an unmapped symbol has no arrow
    assert.deepEqual(finding.examples, ['Webdings F021', 'Wingdings F041', 'Symbol F0E6']);
});

test('a w:sym without w:font uses the run\'s font', async () => {
    const { doc, content, parsed } = await convertBody(p(
        sym(null, 'F0FE', '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/>') + sym(null, 'F0FE')));
    assert.deepEqual(bodyParagraphs(doc, content), ['☑']);
    assert.deepEqual(findingsOf(parsed)[0].examples, ['Wingdings F0FE → ☑', 'Times New Roman F0FE']);
});

test('All Caps upper-cases the run\'s text but not its symbols', async () => {
    const run = `<w:r><w:rPr><w:caps/></w:rPr><w:t xml:space="preserve">açı </w:t>${symEl('Symbol', 'F061')}</w:r>`;
    const { doc, content } = await convertBody(p(run));
    assert.deepEqual(bodyParagraphs(doc, content), ['AÇI α']);
});

test('a symbol in a hidden run writes nothing and is reported as hidden text', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Görünen') + sym('Wingdings', 'F0FE', '<w:vanish/>')));
    assert.deepEqual(bodyParagraphs(doc, content), ['Görünen']);
    assert.equal(content.includes('☑'), false);
    assert.deepEqual(findingsOf(parsed), [{ code: 'hiddenText', count: 1, examples: [] }]);
});

test('symbols in table cells, headers and footers convert the same way', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Gövde')) + tbl([p(sym('Wingdings', 'F0FE') + r(' Evet')), p(sym('Wingdings', 'F0A8') + r(' Hayır'))]), {
            sectPr: references,
            parts: [
                part('header', p(r('Büro ') + sym('Wingdings', 'F0D8'))),
                // A footer that holds only a symbol is still written
                part('footer', p(sym('Symbol', 'F0A7')))
            ]
        });

    const cells = [...doc.querySelectorAll('cell paragraph')]
        .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));
    assert.deepEqual(cells, ['☑ Evet', '□ Hayır']);
    const texts = selector => [...doc.querySelectorAll(`${selector} content[startOffset]`)].map(el => textOf(el, content)).join('');
    assert.equal(texts('header'), 'Büro ➢');
    assert.equal(texts('footer'), '♣');
    assert.ok(content.startsWith('Büro ➢'));
    assert.ok(content.endsWith('♣'));
    assert.equal(findingsOf(parsed)[0].count, 4);
});

test('symbols in referenced footnotes and endnotes are written and counted per reference; unreferenced notes add nothing', async () => {
    const { content, parsed } = await convertBody(
        p(r('Dipnotlu') + '<w:r><w:footnoteReference w:id="1"/></w:r>') +
        p(r('Son notlu') + '<w:r><w:endnoteReference w:id="1"/></w:r>'), {
            footnotes: {
                1: p(sym('Wingdings', 'F0FC') + r(' onaylandı')),
                2: p(sym('Wingdings', 'F0FB') + r(' kullanılmayan'))
            },
            endnotes: { 1: p(r('Seçim ') + sym('Wingdings 2', 'F054')) }
        });

    assert.ok(content.includes('✓ onaylandı'));
    assert.ok(content.includes('Seçim ☒'));
    assert.equal(content.includes('✗'), false);
    const symbols = findingsOf(parsed).find(finding => finding.code === 'symbols');
    assert.deepEqual(symbols, { code: 'symbols', count: 2, examples: ['Wingdings F0FC → ✓', 'Wingdings 2 F054 → ☒'] });
});

test('a positional tab (w:ptab) becomes a tab', async () => {
    const run = '<w:r><w:t>Sol</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t>sağ</w:t></w:r>';
    const { content, parsed } = await convertBody(p(run));
    assert.deepEqual(plain(parsed.elements[0].runs.map(item => item.type)), ['text', 'tab', 'text']);
    assert.ok(content.startsWith('Sol\tsağ'));
});
