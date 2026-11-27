const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

// Paragraph shading (w:shd in w:pPr) and underline styles other than single (w:u w:val). Docs.md gives
// the <paragraph> no background color and the <content> one underline, so neither is written: a shaded
// paragraph has no background and a styled underline is written underline="true". Both are reported,
// as paragraphShading and underlineStyles conversion notes; the UDF itself is unchanged.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const style = (type, id, { pPr = '', rPr = '', basedOn = '' } = {}) =>
    `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${id}"/>` +
    `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:style>`;
const docDefaults = ({ pPr = '', rPr = '' }) => '<w:docDefaults>' +
    (rPr ? `<w:rPrDefault><w:rPr>${rPr}</w:rPr></w:rPrDefault>` : '') +
    (pPr ? `<w:pPrDefault><w:pPr>${pPr}</w:pPr></w:pPrDefault>` : '') + '</w:docDefaults>';
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const footnoteRef = id => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const textBox = inner => '<w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:graphic>' +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const GREY = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>';
const u = val => `<w:u w:val="${val}"/>`;

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;
const examples = (parsed, code) => occurrences(parsed, code).map(o => o.example);

// The attributes of the one <content> whose text is text (offsets left out)
const attrs = (doc, content, text) => {
    const all = [...doc.querySelectorAll('content')].filter(el => textOf(el, content) === text);
    assert.equal(all.length, 1, `one <content> for ${text}`);
    return Object.fromEntries([...all[0].attributes]
        .filter(a => !['startOffset', 'length'].includes(a.name)).map(a => [a.name, a.value]));
};

const PLAIN = { family: 'Times New Roman', size: '12' };

test('the issue\'s reproductions: a grey-shaded paragraph and double, dotted and words underlines each give one note', async () => {
    const { doc, xml, content, parsed } = await convertBody(
        p(r('AÇIKLAMALAR'), GREY) +
        p(r('Çift', u('double')) + r(' noktalı', u('dotted')) + r(' sözcük sözcük', u('words'))));
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'underlineStyles', count: 1, examples: ['Çift noktalı sözcük sözcük'] },
        { code: 'paragraphShading', count: 1, examples: ['AÇIKLAMALAR'] }
    ]);
    assert.deepEqual(occurrences(parsed, 'paragraphShading'),
        [{ example: 'AÇIKLAMALAR', location: { kind: 'paragraph', paragraph: 1 } }]);
    assert.deepEqual(occurrences(parsed, 'underlineStyles'),
        [{ example: 'Çift noktalı sözcük sözcük', location: { kind: 'paragraph', paragraph: 2, heading: 'AÇIKLAMALAR' } }]);

    // Nothing new is written: no paragraph background, and a plain underline
    assert.doesNotMatch(xml, /<paragraph[^>]*(background|[Ss]had|[Ff]ill)/);
    assert.deepEqual(attrs(doc, content, 'AÇIKLAMALAR'), PLAIN);
    for (const text of ['Çift', ' noktalı', ' sözcük sözcük']) {
        assert.deepEqual(attrs(doc, content, text), { ...PLAIN, underline: 'true' }, text);
    }
});

test('every styled underline is reported, and none of them adds an attribute or an internal key', async () => {
    const styled = ['words', 'double', 'thick', 'dotted', 'dottedHeavy', 'dash', 'dashedHeavy', 'dashLong', 'dashLongHeavy',
        'dotDash', 'dashDotHeavy', 'dotDotDash', 'dashDotDotHeavy', 'wave', 'wavyHeavy', 'wavyDouble'];
    const { doc, content, parsed } = await convertBody(styled.map(val => p(r(val, u(val)))).join(''));
    assert.deepEqual(examples(parsed, 'underlineStyles'), styled.slice(0, 5));
    assert.equal(finding(parsed, 'underlineStyles').count, styled.length);
    for (const val of styled) assert.deepEqual(attrs(doc, content, val), { ...PLAIN, underline: 'true' }, val);
    for (const paragraph of plain(parsed.elements)) {
        for (const run of paragraph.runs) assert.equal('underlineStyle' in run, false);
    }
});

test('single and none underlines, and a shading with an auto fill, a nil pattern or no color, give no note', async () => {
    const { doc, xml, content, parsed } = await convertBody(
        p(r('tek', u('single')) + r(' yok', u('none'))) +
        p(r('Otomatik'), '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>') +
        p(r('Nil'), '<w:shd w:val="nil" w:fill="D9D9D9"/>') +
        p(r('Renksiz'), '<w:shd w:val="clear"/>') +
        p(r('Bozuk'), '<w:shd w:val="clear" w:fill="grey"/>'));
    assert.deepEqual(attrs(doc, content, 'tek'), { ...PLAIN, underline: 'true' });
    assert.deepEqual(attrs(doc, content, ' yok'), PLAIN);
    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.doesNotMatch(xml, /<paragraph[^>]*background/);
});

test('shading uses the run shading color rule: the pattern color of a solid pattern, else the fill', async () => {
    const { parsed } = await convertBody(
        p(r('Dolu'), '<w:shd w:val="solid" w:color="000000" w:fill="auto"/>') +
        p(r('Desen'), '<w:shd w:val="pct25" w:color="auto" w:fill="FFFF00"/>') +
        p(r('Dolu otomatik'), '<w:shd w:val="solid" w:color="auto" w:fill="auto"/>'));
    assert.deepEqual(examples(parsed, 'paragraphShading'), ['Dolu', 'Desen']);
});

test('paragraph shading resolves through the paragraph style chain and the document defaults, and nil clears it', async () => {
    const styles = docDefaults({ pPr: GREY }) +
        style('paragraph', 'Normal', {}) +
        style('paragraph', 'Gri', { pPr: GREY.replace('D9D9D9', 'BFBFBF') }) +
        style('paragraph', 'GriAlt', { basedOn: 'Gri' }) +
        style('paragraph', 'Temiz', { pPr: '<w:shd w:val="nil"/>' });
    const { parsed } = await convertBody(
        p(r('Varsayılan')) +
        p(r('Stil'), '<w:pStyle w:val="Gri"/>') +
        p(r('Zincir'), '<w:pStyle w:val="GriAlt"/>') +
        p(r('Stil temiz'), '<w:pStyle w:val="Temiz"/>') +
        p(r('Kendi temiz'), '<w:pStyle w:val="Gri"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/>'),
        { styles });
    assert.deepEqual(examples(parsed, 'paragraphShading'), ['Varsayılan', 'Stil', 'Zincir']);
    assert.equal(finding(parsed, 'paragraphShading').count, 3);
});

test('a styled underline resolves through the document defaults, the paragraph and character style chains and the run', async () => {
    const styles = style('paragraph', 'CiftPara', { rPr: u('double') }) +
        style('paragraph', 'CiftParaAlt', { basedOn: 'CiftPara' }) +
        style('character', 'Dalga', { rPr: u('wave') }) +
        style('character', 'DalgaAlt', { basedOn: 'Dalga' }) +
        style('character', 'Tek', { rPr: u('single') });
    const { doc, content, parsed } = await convertBody(
        p(r('paragraf'), '<w:pStyle w:val="CiftPara"/>') +
        p(r('zincir'), '<w:pStyle w:val="CiftParaAlt"/>') +
        p(r('karakter', '<w:rStyle w:val="DalgaAlt"/>')) +
        // A later layer's single or none underline replaces the inherited style
        p(r('tekle', '<w:rStyle w:val="Tek"/>'), '<w:pStyle w:val="CiftPara"/>') +
        p(r('kapalı', u('none')), '<w:pStyle w:val="CiftPara"/>') +
        p(r('doğrudan tek', `<w:rStyle w:val="Dalga"/>${u('single')}`)),
        { styles });
    assert.deepEqual(examples(parsed, 'underlineStyles'), ['paragraf', 'zincir', 'karakter']);
    assert.equal(finding(parsed, 'underlineStyles').count, 3);
    for (const text of ['paragraf', 'zincir', 'karakter', 'tekle', 'doğrudan tek']) {
        assert.equal(attrs(doc, content, text).underline, 'true', text);
    }
    assert.equal(attrs(doc, content, 'kapalı').underline, undefined);

    const fromDefaults = await convertBody(p(r('varsayılan')) + p(r('düz', u('single'))), { styles: docDefaults({ rPr: u('dotted') }) });
    assert.deepEqual(examples(fromDefaults.parsed, 'underlineStyles'), ['varsayılan']);
});

test('whitespace-only and hidden styled-underline runs add no note', async () => {
    const { content, parsed } = await convertBody(
        p(r('Görünen ') + r('gizli', `<w:vanish/>${u('double')}`) + r('   ', u('double'))) + p(r('\t', u('wave'))));
    assert.equal(content.includes('gizli'), false);
    assert.equal(finding(parsed, 'underlineStyles'), undefined);
});

test('a body paragraph holding only a page break is not written as a paragraph, so its shading is not reported', async () => {
    const { parsed } = await convertBody(
        p(r('Önce')) + p('<w:r><w:br w:type="page"/></w:r>', GREY) + p(r('Sonra')));
    assert.equal(finding(parsed, 'paragraphShading'), undefined);
});

test('both notes are reported in a table cell, a text box, the header and footer, and a footnote', async () => {
    const parts = [
        { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Büro başlığı'), GREY) },
        { name: 'footer1.xml', root: 'ftr', type: 'footer', id: 'rIdFooter1', inner: p(r('Sayfa ') + r('altı', u('double'))) }
    ];
    const sectPr = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';
    const { parsed } = await convertBody(
        tbl([p(r('Hücre'), GREY), p(r('Bedel ') + r('500 TL', u('thick')))]) +
        p(r('Kutu: ') + textBox(p(r('Kutu metni'), GREY) + p(r('Kutu ') + r('çizgi', u('dash'))))) +
        p(r('Dipnot') + footnoteRef(1)),
        { parts, sectPr, footnotes: { 1: p(r('Gölgeli not'), GREY) + p(r('Not ') + r('çizgisi', u('dotted'))) } });

    const shading = occurrences(parsed, 'paragraphShading');
    assert.deepEqual(shading.map(o => o.example).sort(), ['Büro başlığı', 'Gölgeli not', 'Hücre', 'Kutu metni'].sort());
    assert.deepEqual(shading.find(o => o.example === 'Hücre').location, { kind: 'cell', table: 1, row: 1, cell: 1 });
    assert.deepEqual(shading.find(o => o.example === 'Gölgeli not').location,
        { kind: 'paragraph', paragraph: 2, note: { kind: 'footnote' } });

    const underline = occurrences(parsed, 'underlineStyles');
    assert.deepEqual(underline.map(o => o.example).sort(), ['Bedel 500 TL', 'Kutu çizgi', 'Not çizgisi', 'Sayfa altı'].sort());
    assert.deepEqual(underline.find(o => o.example === 'Bedel 500 TL').location, { kind: 'cell', table: 1, row: 1, cell: 2 });
    assert.deepEqual(underline.find(o => o.example === 'Not çizgisi').location,
        { kind: 'paragraph', paragraph: 2, note: { kind: 'footnote' } });
});

test('a footnote\'s notes are counted once per reference, and its hidden styled-underline run adds none', async () => {
    const { parsed } = await convertBody(p(r('Bir') + footnoteRef(1)) + p(r('İki') + footnoteRef(1)),
        { footnotes: { 1: p(r('Not') + r(' gizli', `<w:vanish/>${u('double')}`), GREY) } });
    assert.equal(finding(parsed, 'paragraphShading').count, 2);
    assert.equal(finding(parsed, 'underlineStyles'), undefined);
});

test('an equation in a paragraph whose style underlines it doubly is reported, in the body and a footnote', async () => {
    const oMath = text => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`;
    const styles = style('paragraph', 'CiftPara', { rPr: u('double') });
    const { doc, content, parsed } = await convertBody(
        p(oMath('A/365'), '<w:pStyle w:val="CiftPara"/>') + p(oMath('x+1')) + p(r('Dipnot') + footnoteRef(1)),
        { styles, footnotes: { 1: p(oMath('B/2'), '<w:pStyle w:val="CiftPara"/>') } });
    assert.equal(attrs(doc, content, 'A/365').underline, 'true');
    assert.equal(attrs(doc, content, 'x+1').underline, undefined);
    assert.deepEqual(examples(parsed, 'underlineStyles'), ['A/365', 'B/2']);
});

test('run shading and highlighting are still written as background, and add no paragraphShading note', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('vurgu', '<w:highlight w:val="yellow"/>') + r(' gölge', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>')));
    assert.equal(attrs(doc, content, 'vurgu').background, '-256');
    assert.ok(attrs(doc, content, ' gölge').background);
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('the petition fixtures give neither note', async () => {
    for (const petition of petitions) {
        const { DocxParser } = require('./helpers').loadConverter();
        const parsed = await new DocxParser().parse(await petition.build());
        const codes = plain(parsed.findings).map(f => f.code);
        assert.ok(!codes.includes('paragraphShading') && !codes.includes('underlineStyles'), petition.name);
    }
});
