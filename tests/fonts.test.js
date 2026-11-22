const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx, declareSize, convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const rFonts = attributes => `<w:rFonts ${attributes}/>`;
const theme = (slot, kind = 'HAnsi') => rFonts(`w:asciiTheme="${slot}${kind}" w:hAnsiTheme="${slot}HAnsi" ` +
    `w:eastAsiaTheme="${slot}EastAsia" w:cstheme="${slot}Bidi"`);
const style = (type, id, rPr, more = '') =>
    `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/>${more}<w:rPr>${rPr}</w:rPr></w:style>`;
const docDefaults = rPr => `<w:docDefaults><w:rPrDefault><w:rPr>${rPr}</w:rPr></w:rPrDefault></w:docDefaults>`;
const pStyle = id => `<w:pStyle w:val="${id}"/>`;

// The fonts of Word's default template before Office 2023
const CALIBRI = { major: 'Calibri Light', minor: 'Calibri' };
// What Word's Normal template writes: theme fonts in w:docDefaults, the major one in the heading styles
const WORD_STYLES = docDefaults(theme('minor') + '<w:sz w:val="22"/>') +
    style('paragraph', 'Normal', '') +
    style('paragraph', 'Heading1', theme('major') + '<w:sz w:val="32"/>', '<w:basedOn w:val="Normal"/>');

/**
 * The font family of every content element, in the body, a table, the header or anywhere
 * else in content.xml, that covers exactly this text
 */
function familiesOf({ doc, content }, text) {
    const families = [...doc.querySelectorAll('elements [startOffset]')]
        .filter(el => textOf(el, content) === text)
        .map(el => el.getAttribute('family'));
    assert.ok(families.length > 0, `no content element covers "${text}"`);
    return families;
}

// The parser logs a theme part it can't read, and the error it rethrows; keep the test output clean
test.beforeEach(t => {
    t.mock.method(console, 'warn', () => {});
    t.mock.method(console, 'error', () => {});
});

test('theme fonts in docDefaults and a heading style resolve through the theme part', async () => {
    const result = await convertBody(p(r('Dilekçe metni')) + p(r('SONUÇ VE İSTEM'), pStyle('Heading1')),
        { styles: WORD_STYLES, theme: CALIBRI });
    assert.deepEqual(familiesOf(result, 'Dilekçe metni'), ['Calibri']);
    assert.deepEqual(familiesOf(result, 'SONUÇ VE İSTEM'), ['Calibri Light']);
    assert.deepEqual(findingSummaries(result.parsed.findings), []);
});

test('minorAscii and majorAscii resolve the same as minorHAnsi and majorHAnsi', async () => {
    const result = await convertBody(p(r('Küçük', theme('minor', 'Ascii'))) + p(r('Büyük', theme('major', 'Ascii'))),
        { theme: CALIBRI });
    assert.deepEqual(familiesOf(result, 'Küçük'), ['Calibri']);
    assert.deepEqual(familiesOf(result, 'Büyük'), ['Calibri Light']);
});

test('the most specific layer decides; within one w:rFonts the theme font wins over w:ascii', async () => {
    const styles = docDefaults(theme('minor')) +
        style('paragraph', 'Explicit', rFonts('w:ascii="Times New Roman" w:hAnsi="Times New Roman"')) +
        style('character', 'Emphasis', theme('major'));
    const result = await convertBody(
        p(r('Açık Arial', rFonts('w:ascii="Arial" w:hAnsi="Arial"'))) +
        p(r('İkisi birden', rFonts('w:ascii="Arial" w:hAnsi="Arial" w:asciiTheme="minorHAnsi"'))) +
        p(r('Stilden', '<w:rStyle w:val="Emphasis"/>'), pStyle('Explicit')) +
        p(r('Paragraf stili'), pStyle('Explicit')),
        { styles, theme: CALIBRI });
    assert.deepEqual(familiesOf(result, 'Açık Arial'), ['Arial']);
    assert.deepEqual(familiesOf(result, 'İkisi birden'), ['Calibri']);
    assert.deepEqual(familiesOf(result, 'Stilden'), ['Calibri Light']);
    assert.deepEqual(familiesOf(result, 'Paragraf stili'), ['Times New Roman']);
});

test('a style\'s theme font overrides a w:ascii its paragraph style chain inherits', async () => {
    const styles = style('paragraph', 'Base', rFonts('w:ascii="Arial"')) +
        style('paragraph', 'Themed', theme('major'), '<w:basedOn w:val="Base"/>');
    const result = await convertBody(p(r('Tema'), pStyle('Themed')) + p(r('Temel'), pStyle('Base')), { styles, theme: CALIBRI });
    assert.deepEqual(familiesOf(result, 'Tema'), ['Calibri Light']);
    assert.deepEqual(familiesOf(result, 'Temel'), ['Arial']);
});

test('a theme font that can\'t be resolved falls back to Times New Roman, or the same w:rFonts\' w:ascii', async () => {
    const body = p(r('Metin')) + p(r('Yedek', rFonts('w:ascii="Arial" w:asciiTheme="majorHAnsi"')));
    const cases = {
        'no theme part': undefined,
        'an empty typeface': { major: '', minor: '' },
        'a blank typeface': { major: '  ', minor: ' ' },
        'no a:latin': {},
        'a malformed part': '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>',
        'a part that is not a theme': '<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    };
    for (const [name, themeOption] of Object.entries(cases)) {
        const result = await convertBody(body, { styles: docDefaults(theme('minor')), theme: themeOption });
        assert.deepEqual(familiesOf(result, 'Metin'), ['Times New Roman'], name);
        assert.deepEqual(familiesOf(result, 'Yedek'), ['Arial'], name);
        assert.deepEqual(findingSummaries(result.parsed.findings), [], name);
    }
});

test('only the slot the theme gives resolves; East Asian and complex script slots don\'t set the Latin font', async () => {
    const result = await convertBody(p(r('Küçük')) + p(r('Büyük', theme('major'))) +
        p(r('Doğu', rFonts('w:ascii="Arial" w:asciiTheme="minorEastAsia"'))),
        { styles: docDefaults(theme('minor')), theme: { minor: 'Aptos' } });
    assert.deepEqual(familiesOf(result, 'Küçük'), ['Aptos']);
    assert.deepEqual(familiesOf(result, 'Büyük'), ['Aptos']);
    assert.deepEqual(familiesOf(result, 'Doğu'), ['Arial']);
});

test('the theme font applies in a table cell, a text box, the header and a footnote', async () => {
    const cellTable = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(r('Hücrede'))}</w:tc></w:tr></w:tbl>`;
    const textBox = '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        `<wps:wsp><wps:txbx><w:txbxContent>${p(r('Kutuda'))}</w:txbxContent></wps:txbx></wps:wsp>` +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    const footnoteRef = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
    const footnoteMark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' + r(' ');
    const result = await convertBody(p(r('Gövdede') + footnoteRef) + cellTable + p(textBox), {
        styles: WORD_STYLES,
        theme: CALIBRI,
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>',
        parts: [{ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Üstte')) }],
        footnotes: { 1: p(footnoteMark + r('Dipnotta')) }
    });
    assert.ok(result.doc.querySelector('elements > header'), 'the header is converted');
    assert.ok(result.doc.querySelector('elements table cell'), 'the table is converted');
    for (const text of ['Gövdede', 'Hücrede', 'Kutuda', 'Üstte', 'Dipnotta']) {
        assert.deepEqual(familiesOf(result, text), ['Calibri'], text);
    }
});

test('documents without theme fonts don\'t read the theme into their run fonts', async () => {
    const body = p(r('Varsayılan')) + p(r('Arial', rFonts('w:ascii="Arial"')));
    const without = await convertBody(body);
    const withTheme = await convertBody(body, { theme: CALIBRI });
    assert.equal(withTheme.xml, without.xml);
    assert.deepEqual(familiesOf(withTheme, 'Varsayılan'), ['Times New Roman']);
});

test('the theme part is read within the decompression limits', async () => {
    const docx = await buildDocx(p(r('Metin')), { styles: docDefaults(theme('minor')), theme: CALIBRI });
    const part = 'word/theme/theme1.xml';
    const parser = new (loadConverter().DocxParser)();
    await assert.rejects(parser.parse(declareSize(docx, part, 0xFFFFFFFE)), error => error.code === 'tooLarge');

    // The theme part counts toward the total: with one byte less than every part read, the parse is refused
    const zip = await JSZip.loadAsync(docx);
    const total = Object.values(zip.files).filter(file => !file.dir && file.name !== '[Content_Types].xml')
        .reduce((sum, file) => sum + file._data.uncompressedSize, 0);
    parser.limits = { part: 100 * 1024 * 1024, total };
    assert.equal((await parser.parse(docx)).elements.length, 1);
    parser.limits = { part: 100 * 1024 * 1024, total: total - 1 };
    await assert.rejects(parser.parse(docx), error => error.code === 'tooLarge');
});
