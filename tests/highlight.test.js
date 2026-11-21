const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraphStyle = (id, rPr) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;
const characterStyle = (id, rPr) => `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const header = inner => ({ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner });
const headerRef = '<w:headerReference w:type="default" r:id="rIdHeader1"/>';

/** The Java-style signed ARGB integer the UDF uses for an opaque RGB hex color */
const argb = hex => (0xFF000000 | parseInt(hex, 16)) | 0;

/** [text, background attribute or null] for each <content> under `root` */
const backgrounds = (root, content) => [...root.querySelectorAll('content')]
    .map(el => [textOf(el, content), el.getAttribute('background')]);

test('a highlighted run writes its color as the content background; its neighbours write none', async () => {
    const { doc, content, xml, parsed } = await convertBody(
        p(r('Talep edilen tutar: ') + r('125.000 TL', '<w:highlight w:val="yellow"/>') + r(' olup')));

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content),
        [['Talep edilen tutar: ', null], ['125.000 TL', '-256'], [' olup', null]]);
    assert.equal(argb('FFFF00'), -256);
    // No run carries the helper properties the background is resolved from
    for (const run of parsed.elements[0].runs) {
        for (const key of ['highlight', 'shading']) assert.equal(key in run, false, key);
    }
    // background is the only attribute added
    assert.match(xml, /<content startOffset="\d+" length="10" family="Times New Roman" size="12" background="-256" \/>/);
});

test('every Word highlight color maps to its palette value', async () => {
    const palette = {
        yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
        darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
        darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF'
    };
    const names = Object.keys(palette);
    const { doc, content } = await convertBody(
        p(names.map(name => r(name, `<w:highlight w:val="${name}"/>`)).join('')));

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content),
        names.map(name => [name, String(argb(palette[name]))]));
});

test('"none" or an unknown highlight value writes no background', async () => {
    const { doc, content } = await convertBody(
        p(r('yok', '<w:highlight w:val="none"/>') + r('bilinmeyen', '<w:highlight w:val="orange"/>') +
          r('boş', '<w:highlight/>')));

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content),
        [['yok', null], ['bilinmeyen', null], ['boş', null]]);
});

test('run shading writes its fill, a solid pattern its pattern color, and nil or auto shading none', async () => {
    const { doc, content } = await convertBody(p(
        r('gri', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>') +
        r('desen', '<w:shd w:val="pct25" w:color="000000" w:fill="c6d9f1"/>') +
        r('otomatik', '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>') +
        r('dolgusuz', '<w:shd w:val="clear" w:color="auto"/>') +
        r('nil', '<w:shd w:val="nil" w:fill="D9D9D9"/>') +
        r('kırmızı', '<w:shd w:val="solid" w:color="FF0000" w:fill="auto"/>') +
        r('düz-otomatik', '<w:shd w:val="solid" w:color="auto" w:fill="00FF00"/>') +
        r('bozuk', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9"/>')));

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content), [
        ['gri', String(argb('D9D9D9'))],
        ['desen', String(argb('C6D9F1'))],
        ['otomatik', null],
        ['dolgusuz', null],
        ['nil', null],
        ['kırmızı', String(argb('FF0000'))],
        ['düz-otomatik', String(argb('00FF00'))],
        ['bozuk', null]
    ]);
});

test('a run with both highlighting and shading writes the highlight color', async () => {
    const { doc, content } = await convertBody(p(
        r('ikisi', '<w:highlight w:val="green"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>')));

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content), [['ikisi', String(argb('00FF00'))]]);
});

test('highlighting and shading resolve through the styles, and the run can turn them off', async () => {
    const styles =
        characterStyle('Vurgu', '<w:highlight w:val="cyan"/>') +
        characterStyle('Gölge', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>') +
        paragraphStyle('Önemli', '<w:highlight w:val="yellow"/>') +
        paragraphStyle('Gölgeli', '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>');
    const { doc, content } = await convertBody(
        p(r('karakter', '<w:rStyle w:val="Vurgu"/>') + r(' gölge', '<w:rStyle w:val="Gölge"/>')) +
        p(r('paragraf') + r(' kapalı', '<w:highlight w:val="none"/>') +
          r(' karakter üstte', '<w:rStyle w:val="Vurgu"/>'), '<w:pStyle w:val="Önemli"/>') +
        p(r('gölgeli') + r(' vurgulu', '<w:highlight w:val="red"/>') +
          r(' gölgesiz', '<w:shd w:val="nil"/>'), '<w:pStyle w:val="Gölgeli"/>') +
        // The style's highlight is turned off, but its shading still shows under it
        p(r('altta gölge', '<w:highlight w:val="none"/>'), '<w:pStyle w:val="Gölgeli"/>'),
        { styles });

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content), [
        ['karakter', String(argb('00FFFF'))],
        [' gölge', String(argb('D9D9D9'))],
        ['paragraf', '-256'],
        [' kapalı', null],
        [' karakter üstte', String(argb('00FFFF'))],
        ['gölgeli', String(argb('F2F2F2'))],
        [' vurgulu', String(argb('FF0000'))],
        [' gölgesiz', null],
        ['altta gölge', String(argb('F2F2F2'))]
    ]);
});

test('highlighting set in the document defaults applies, and a highlight under shading survives shading turned off', async () => {
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:highlight w:val="lightGray"/></w:rPr></w:rPrDefault></w:docDefaults>' +
        characterStyle('Gölge', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>');
    const { doc, content } = await convertBody(
        p(r('varsayılan') + r(' gölgeli', '<w:rStyle w:val="Gölge"/>') + r(' kapalı', '<w:highlight w:val="none"/>')),
        { styles });

    assert.deepEqual(backgrounds(doc.querySelector('elements'), content), [
        ['varsayılan', String(argb('C0C0C0'))],
        [' gölgeli', String(argb('C0C0C0'))],
        [' kapalı', null]
    ]);
});

test('highlighted text in a table cell and in the header writes a background too', async () => {
    const { doc, content } = await convertBody(
        tbl([p(r('Hücre ') + r('süre', '<w:highlight w:val="yellow"/>')),
            p(r('gri', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'))]) + p(r('Gövde')),
        {
            sectPr: headerRef,
            parts: [header(p(r('Esas No: ') + r('2026/123', '<w:highlight w:val="green"/>')))]
        });

    const cells = [...doc.querySelectorAll('cell')];
    assert.deepEqual(backgrounds(cells[0], content), [['Hücre ', null], ['süre', '-256']]);
    assert.deepEqual(backgrounds(cells[1], content), [['gri', String(argb('D9D9D9'))]]);
    assert.deepEqual(backgrounds(doc.querySelector('elements > header'), content),
        [['Esas No: ', null], ['2026/123', String(argb('00FF00'))]]);
});

test('a document without highlighting or run shading writes no background attribute', async () => {
    const { xml } = await convertBody(
        p(r('Düz') + r(' renkli', '<w:color w:val="C00000"/>') + r(' kalın', '<w:b/>')) +
        tbl([p(r('Hücre'))]) +
        // Paragraph shading is not converted: the documented <paragraph> has no background
        p(r('Paragraf gölgesi'), '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'));

    assert.equal(xml.includes('background='), false);
});
