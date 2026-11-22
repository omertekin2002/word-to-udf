const { test } = require('node:test');
const assert = require('node:assert/strict');
const { buildDocx, convert, convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const HEADER_TEXT = 'Ad Soyad Hukuk Bürosu';
const FOOTER_TEXT = 'Dosya No: 2026/0000';

/**
 * A header (w:hdr) or footer (w:ftr) part for buildDocx, related as rIdHeader<n> / rIdFooter<n>,
 * with its own image relationships (id -> target relative to the part), if any
 */
const part = (kind, inner, n = 1, rels = undefined) => ({
    name: `${kind}${n}.xml`,
    root: kind === 'header' ? 'hdr' : 'ftr',
    type: kind,
    id: `rId${kind === 'header' ? 'Header' : 'Footer'}${n}`,
    inner,
    rels
});
const reference = (kind, n = 1, type = 'default') =>
    `<w:${kind}Reference w:type="${type}" r:id="rId${kind === 'header' ? 'Header' : 'Footer'}${n}"/>`;
const bothReferences = reference('header') + reference('footer');

// A complex field (w:fldChar) and a simple field (w:fldSimple), each showing a stale result
const complexField = (instruction, result) =>
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:r><w:t>${result}</w:t></w:r>` +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const simpleField = (instruction, result) => `<w:fldSimple w:instr=" ${instruction} ">${r(result)}</w:fldSimple>`;

const PIC_GRAPHIC = rId => '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>`;
/** A run holding an inline DrawingML picture, 1 in by 0.5 in (72 by 36 points) */
const pictureRun = (rId = 'rIdLogo') =>
    `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>${PIC_GRAPHIC(rId)}</wp:inline></w:drawing></w:r>`;
/** A run holding a picture anchored to the page, behind or beside the text, 2 in by 1 in */
const anchoredPictureRun = (rId = 'rIdLogo') => '<w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="1">' +
    '<wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:posOffset>228600</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="1828800" cy="914400"/><wp:wrapNone/>${PIC_GRAPHIC(rId)}</wp:anchor></w:drawing></w:r>`;
/** A run holding a VML picture (w:pict), 120 by 40 points */
const vmlPictureRun = (rId = 'rIdLogo') => '<w:r><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:120pt;height:40pt">' +
    `<v:imagedata r:id="${rId}"/></v:shape></w:pict></w:r>`;
/** Word's way of writing a picture for older readers: DrawingML, with a VML fallback of the same picture */
const alternatePictureRun = (rId = 'rIdLogo') => '<w:r><mc:AlternateContent>' +
    `<mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>${PIC_GRAPHIC(rId)}</wp:inline></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="${rId}"/></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';
const chartRun = '<w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
/** Word's picture watermark (Design › Watermark): a VML shape in the header behind the page */
const watermarkRun = (rId = 'rIdLogo') => '<w:r><w:pict><v:shape id="WordPictureWatermark1" type="#_x0000_t75" ' +
    'style="position:absolute;width:415pt;height:415pt;z-index:-251657216;mso-position-horizontal:center;mso-position-vertical:center">' +
    `<v:imagedata r:id="${rId}" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r>`;
const picture = pictureRun();

const PNG = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x01]);
const OTHER_PNG = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x02]);
const GIF = Buffer.from('GIF89a\x01\x00\x01\x00', 'latin1');
const LOGO = { 'word/media/logo.png': PNG };
const LOGO_RELS = { rIdLogo: 'media/logo.png' };
const textBox = '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData>' +
    `<wps:wsp><wps:txbx><w:txbxContent>${p(r('Kutu metni'))}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${p(r('Kutu metni'))}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';
const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc>${p(r('Tablo hücresi'))}</w:tc></w:tr></w:tbl>`;

const plain = value => JSON.parse(JSON.stringify(value));

/** The element children of <elements>, by local name */
const topLevel = doc => [...doc.querySelector('elements').children];

/** Text of every content/tab element under a root, from the CDATA pool */
const texts = (root, content) => [...root.querySelectorAll('[startOffset]')].map(el => textOf(el, content));

/** Every <content>/<tab> of a document, in document order, with its section */
function pieces(doc, content) {
    return [...doc.querySelectorAll('elements [startOffset]')].map(el => ({
        section: el.closest('header') ? 'header' : el.closest('footer') ? 'footer' : 'body',
        start: Number(el.getAttribute('startOffset')),
        text: textOf(el, content)
    }));
}

test('a default header and footer are written as the first and last elements, each holding its paragraph', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Birinci paragraf')) + p(r('İkinci paragraf')), {
        sectPr: bothReferences,
        parts: [part('header', p(r(HEADER_TEXT))), part('footer', p(r(FOOTER_TEXT)))]
    });

    const children = topLevel(doc);
    assert.deepEqual(children.map(el => el.localName), ['header', 'paragraph', 'paragraph', 'footer']);
    const [header, , , footer] = children;
    assert.deepEqual([...header.children].map(el => el.localName), ['paragraph']);
    assert.deepEqual([...footer.children].map(el => el.localName), ['paragraph']);
    assert.deepEqual(texts(header, content), [HEADER_TEXT]);
    assert.deepEqual(texts(footer, content), [FOOTER_TEXT]);

    // The pool is header text, then the body, then footer text
    assert.equal(content, HEADER_TEXT + 'Birinci paragraf' + 'İkinci paragraf' + FOOTER_TEXT);
    assert.ok(content.startsWith(HEADER_TEXT));
    assert.ok(content.endsWith(FOOTER_TEXT));

    // No page number attributes: their format isn't documented
    assert.equal(footer.attributes.length, 0);
    assert.equal(header.attributes.length, 0);

    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.equal(parsed.header.length, 1);
    assert.equal(parsed.footer.length, 1);
});

test('every offset in header, body and footer points at its own text, in pool order', async () => {
    const { doc, content } = await convertBody(p(r('Gövde', '<w:b/>') + r(' metni')) + p(r('Son paragraf')), {
        sectPr: bothReferences,
        parts: [
            part('header', p(r('Ad Soyad', '<w:b/>') + r(' Hukuk Bürosu')) + p(r('Örnek Mah. 1. Sok. No: 0'))),
            part('footer', p(r('Dosya No: ') + r('2026/0000', '<w:i/>')))
        ]
    });

    const found = pieces(doc, content);
    assert.deepEqual(found.map(({ section, text }) => [section, text]), [
        ['header', 'Ad Soyad'], ['header', ' Hukuk Bürosu'], ['header', 'Örnek Mah. 1. Sok. No: 0'],
        ['body', 'Gövde'], ['body', ' metni'], ['body', 'Son paragraf'],
        ['footer', 'Dosya No: '], ['footer', '2026/0000']
    ]);
    // Contiguous and in order: each piece starts where the one before it ends, covering the whole pool
    let offset = 0;
    for (const { start, text } of found) {
        assert.equal(start, offset);
        offset += text.length;
    }
    assert.equal(offset, content.length);
});

test('the footer text comes after appended footnotes and endnotes', async () => {
    const { doc, content } = await convertBody(
        p(r('Dipnotlu metin') + '<w:r><w:footnoteReference w:id="1"/></w:r>' +
            r(' ve son notlu metin') + '<w:r><w:endnoteReference w:id="1"/></w:r>'), {
            sectPr: bothReferences,
            footnotes: { 1: p(r('Dipnot metni')) },
            endnotes: { 1: p(r('Son not metni')) },
            parts: [part('header', p(r(HEADER_TEXT))), part('footer', p(r(FOOTER_TEXT)))]
        });

    const children = topLevel(doc);
    assert.equal(children[0].localName, 'header');
    assert.equal(children.at(-1).localName, 'footer');
    assert.ok(children.slice(1, -1).every(el => el.localName === 'paragraph'));
    assert.ok(content.startsWith(HEADER_TEXT + 'Dipnotlu metin'));
    assert.ok(content.endsWith('i. Son not metni' + FOOTER_TEXT), JSON.stringify(content));
    assert.ok(content.indexOf('1. Dipnot metni') < content.indexOf('i. Son not metni'));
    assert.deepEqual(texts(children.at(-1), content), [FOOTER_TEXT]);
});

test('header and footer paragraphs resolve styles, fonts, emphasis and alignment like body paragraphs', async () => {
    const styles =
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/>' +
        '<w:pPr><w:jc w:val="center"/><w:tabs><w:tab w:val="center" w:pos="4536"/><w:tab w:val="right" w:pos="9072"/></w:tabs></w:pPr>' +
        '<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr></w:style>' +
        '<w:style w:type="character" w:styleId="Vurgu"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>';
    const paragraph = text => p(r(text, '<w:b/>') + '<w:r><w:tab/></w:r>' + r(' ek', '<w:rStyle w:val="Vurgu"/><w:sz w:val="18"/>'),
        '<w:pStyle w:val="Header"/><w:ind w:left="360" w:right="240"/><w:spacing w:before="120" w:after="60" w:line="360" w:lineRule="auto"/>');

    const { doc, content } = await convertBody(paragraph('Gövdede'), {
        styles,
        sectPr: bothReferences,
        parts: [part('header', paragraph('Üstte')), part('footer', paragraph('Altta'))]
    });

    // Strip offsets and lengths: everything else about the paragraph and its runs must match the body's
    const shape = paragraphEl => paragraphEl.outerHTML.replace(/ (startOffset|length)="\d+"/g, '');
    const [header, body, footer] = topLevel(doc);
    const bodyShape = shape(body);
    assert.equal(shape(header.firstElementChild), bodyShape);
    assert.equal(shape(footer.firstElementChild), bodyShape);

    assert.equal(body.getAttribute('Alignment'), '1');
    assert.equal(body.getAttribute('TabSet'), '226.8:1:0,453.6:2:0');
    const runs = header.querySelectorAll('content');
    assert.equal(textOf(runs[0], content), 'Üstte');
    assert.equal(runs[0].getAttribute('family'), 'Arial');
    assert.equal(runs[0].getAttribute('size'), '10');
    assert.equal(runs[0].getAttribute('bold'), 'true');
    assert.equal(runs[1].getAttribute('italic'), 'true');
    assert.equal(runs[1].getAttribute('size'), '9');
    assert.equal(header.querySelectorAll('tab').length, 1);
});

test('fields are left out of the converted footer, their results too, and noted', async () => {
    // Word's page number footer: a content control around the paragraph, with PAGE and NUMPAGES fields
    const footer = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Page Numbers (Bottom of Page)"/></w:docPartObj></w:sdtPr>' +
        `<w:sdtContent>${p(r(FOOTER_TEXT + ' — Sayfa ') + complexField('PAGE', '7') + r(' / ') + simpleField('NUMPAGES', '9'))}</w:sdtContent></w:sdt>`;
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('footer'),
        parts: [part('footer', footer)]
    });

    const footerEl = topLevel(doc).at(-1);
    assert.equal(footerEl.localName, 'footer');
    assert.deepEqual(texts(footerEl, content), [FOOTER_TEXT + ' — Sayfa ', ' / ']);
    assert.ok(!/[79]/.test(content), 'no stale field result');
    assert.ok(!content.includes('PAGE'));
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a field spanning paragraphs and nested fields are left out up to their matching end', async () => {
    const nested = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF </w:instrText></w:r>' +
        complexField('PAGE', '1') + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r('dış sonuç');
    const header = p(r('Önce ') + nested) + p(r('hâlâ alanda') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + r(' Sonra'));
    const { doc, content } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', header)]
    });

    const headerEl = topLevel(doc)[0];
    assert.deepEqual(texts(headerEl, content), ['Önce ', ' Sonra']);
});

test('text sharing a run with a field boundary is kept; only the field is left out', async () => {
    const footer = p('<w:r><w:t xml:space="preserve">Önce </w:t><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>4</w:t></w:r>' +
        `<w:r><w:fldChar w:fldCharType="end"/><w:t xml:space="preserve"> ${FOOTER_TEXT}</w:t></w:r>`);
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('footer'),
        parts: [part('footer', footer)]
    });
    assert.deepEqual(texts(topLevel(doc).at(-1), content), ['Önce ', ` ${FOOTER_TEXT}`]);
    assert.ok(!content.includes('4'));
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a footer holding only a page number field writes no footer, and is noted', async () => {
    const { doc, xml, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('footer'),
        parts: [part('footer', p(complexField('PAGE', '1')))]
    });
    assert.deepEqual(topLevel(doc).map(el => el.localName), ['paragraph']);
    assert.ok(!xml.includes('<footer'));
    assert.equal(parsed.footer, null);
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

for (const [name, inner] of [['a text box', p(r(HEADER_TEXT) + textBox)], ['a table', table + p(r(HEADER_TEXT))],
    ['a table holding a picture', table.replace('Tablo hücresi</w:t></w:r>', `Tablo hücresi</w:t></w:r>${picture}`) + p(r(HEADER_TEXT))]]) {
    test(`${name} in the converted header is left out and noted; the header's paragraphs are kept`, async () => {
        const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
            sectPr: reference('header'),
            parts: [part('header', inner, 1, LOGO_RELS)],
            files: LOGO
        });
        const headerEl = topLevel(doc)[0];
        assert.equal(headerEl.localName, 'header');
        assert.deepEqual(texts(headerEl, content), [HEADER_TEXT]);
        assert.equal(doc.querySelectorAll('image').length, 0);
        assert.ok(!content.includes('Kutu metni') && !content.includes('Tablo hücresi'));
        assert.deepEqual(plain(parsed.images), {}, 'a picture left out is not decoded');
        assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
    });
}

test('tabs and line breaks in a header are kept; a page break is not', async () => {
    const { doc, content } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', p(r('Sol') + '<w:r><w:tab/></w:r>' + r('Sağ') + '<w:r><w:br/></w:r>' + r('Alt satır') +
            '<w:r><w:br w:type="page"/></w:r>'))]
    });
    const headerEl = topLevel(doc)[0];
    assert.deepEqual(texts(headerEl, content), ['Sol', '\t', 'Sağ', '\n', 'Alt satır']);
    assert.equal(doc.querySelectorAll('page-break').length, 0);
    assert.ok(content.startsWith('Sol\tSağ\nAlt satır'));
});

test('header text with CDATA and markup characters stays well-formed and unchanged', async () => {
    const tricky = 'A &amp; B &lt;Hukuk&gt; ]]&gt; "son"';
    // convert() asserts content.xml is well-formed
    const { content, doc } = await convertBody(p(r('Gövde ]]&gt; metni')), {
        sectPr: bothReferences,
        parts: [part('header', p(r(tricky))), part('footer', p(r('Alt ]]&gt;')))]
    });
    assert.equal(content, 'A & B <Hukuk> ]]> "son"' + 'Gövde ]]> metni' + 'Alt ]]>');
    assert.deepEqual(texts(topLevel(doc)[0], content), ['A & B <Hukuk> ]]> "son"']);
});

test('an empty or whitespace-only default header and footer write nothing and change nothing', async () => {
    const body = p(r('Gövde'));
    const without = await convertBody(body);
    const withEmpty = await convertBody(body, {
        sectPr: bothReferences,
        parts: [part('header', p('') + p(r('   '))), part('footer', p('<w:r><w:tab/></w:r>'), 1)]
    });
    assert.equal(withEmpty.xml, without.xml);
    assert.equal(withEmpty.parsed.header, null);
    assert.equal(withEmpty.parsed.footer, null);
    assert.deepEqual(findingSummaries(withEmpty.parsed.findings), []);
});

test('a document with only a header still gets a body paragraph, between header and footer', async () => {
    const { doc, content } = await convertBody('', {
        sectPr: bothReferences,
        parts: [part('header', p(r(HEADER_TEXT))), part('footer', p(r(FOOTER_TEXT)))]
    });
    const children = topLevel(doc);
    assert.deepEqual(children.map(el => el.localName), ['header', 'paragraph', 'footer']);
    assert.equal(content, HEADER_TEXT + '​' + FOOTER_TEXT);
    assert.deepEqual(texts(children[1], content), ['​']);
    assert.deepEqual(texts(children[2], content), [FOOTER_TEXT]);
});

test('the final section uses the previous section\'s default header when it has none of its own', async () => {
    const sectionBreak = p(`<w:pPr><w:sectPr>${reference('header')}</w:sectPr></w:pPr>` + r('Bölüm 1'));
    const { doc, content, parsed } = await convertBody(sectionBreak + p(r('Bölüm 2')), {
        sectPr: '<w:pgSz w:w="11906" w:h="16838"/>',
        parts: [part('header', p(r(HEADER_TEXT)))]
    });
    assert.equal(topLevel(doc)[0].localName, 'header');
    assert.ok(content.startsWith(HEADER_TEXT + 'Bölüm 1'));
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('the final section\'s own header wins over an earlier section\'s, which is noted', async () => {
    const sectionBreak = p(`<w:pPr><w:sectPr>${reference('header', 2)}</w:sectPr></w:pPr>` + r('Bölüm 1'));
    const { doc, content, parsed } = await convertBody(sectionBreak + p(r('Bölüm 2')), {
        sectPr: reference('header'),
        parts: [part('header', p(r(HEADER_TEXT))), part('header', p(r('Eski bölüm başlığı')), 2)]
    });
    assert.deepEqual(texts(topLevel(doc)[0], content), [HEADER_TEXT]);
    assert.ok(!content.includes('Eski bölüm başlığı'));
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('first-page and even-page headers are not converted; the default one is', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header', 1) + reference('header', 2, 'first') + reference('footer', 3, 'even') + '<w:titlePg/>',
        parts: [part('header', p(r(HEADER_TEXT))), part('header', p(r('İlk sayfa')), 2), part('footer', p(r('Çift sayfa')), 3)]
    });
    assert.deepEqual(topLevel(doc).map(el => el.localName), ['header', 'paragraph']);
    assert.equal(content, HEADER_TEXT + 'Gövde');
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a missing or malformed header part writes nothing', async () => {
    const missing = await convertBody(p(r('Gövde')), { sectPr: reference('header') });
    assert.equal(missing.parsed.header, null);
    assert.ok(!missing.xml.includes('<header'));

    const docx = await buildDocx(p(r('Gövde')), { sectPr: reference('header'), parts: [part('header', '<w:p>')] });
    const malformed = await convert(docx);
    assert.equal(malformed.parsed.header, null);
    assert.ok(!malformed.xml.includes('<header'));
});

/** The <image> elements under a root, as plain data */
const imagesIn = root => [...root.querySelectorAll('image')].map(image => ({
    data: image.getAttribute('imageData'),
    width: image.getAttribute('width'),
    height: image.getAttribute('height'),
    length: image.getAttribute('length')
}));
const logo = (width, height) => ({ data: PNG.toString('base64'), width: String(width), height: String(height), length: '1' });

for (const [name, run, size] of [['inline', pictureRun(), [72, 36]], ['anchored', anchoredPictureRun(), [144, 72]]]) {
    test(`an ${name} logo in the default header, related through the header's own rels, is written as an <image> in it`, async () => {
        const { doc, content, xml, parsed } = await convertBody(p(r('Gövde')), {
            sectPr: reference('header'),
            parts: [part('header', p(run + r(HEADER_TEXT)), 1, LOGO_RELS)],
            files: LOGO
        });
        const [headerEl] = topLevel(doc);
        assert.equal(headerEl.localName, 'header');
        assert.deepEqual([...headerEl.children].map(el => el.localName), ['paragraph']);
        assert.deepEqual([...headerEl.firstElementChild.children].map(el => el.localName), ['image', 'content']);
        assert.deepEqual(imagesIn(headerEl), [logo(...size)]);
        // The documented shape: <image imageData width height startOffset length="1"> over a U+FFFC
        const image = headerEl.querySelector('image');
        assert.equal(image.getAttribute('startOffset'), '0');
        assert.equal(textOf(image, content), '￼');
        assert.equal(content, '￼' + HEADER_TEXT + 'Gövde');
        assert.match(xml, /<header><paragraph [^>]*><image startOffset="0" length="1" imageData="[^"]+" width="\d+" height="\d+" \/>/);
        assert.deepEqual(findingSummaries(parsed.findings), []);
    });
}

test('a header picture resolves through the header part\'s rels, not a document relationship with the same id', async () => {
    // The document relates rIdLogo to another picture; the header's own rels point rIdLogo at logo.png
    const docx = await buildDocx(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', p(picture), 1, LOGO_RELS)],
        files: { ...LOGO, 'word/media/other.png': OTHER_PNG }
    });
    // Add the clashing id to the document's own relationships
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(docx);
    const rels = await zip.file('word/_rels/document.xml.rels').async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
        '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="media/other.png"/></Relationships>'));
    const { doc, parsed } = await convert(await zip.generateAsync({ type: 'nodebuffer' }));

    assert.deepEqual(imagesIn(topLevel(doc)[0]), [logo(72, 36)]);
    assert.deepEqual(Object.keys(plain(parsed.images)), ['word/media/logo.png']);
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('a header holding only a logo is written; the pool starts with U+FFFC and every offset still points at its own text', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Gövde', '<w:b/>') + r(' metni')) + p(r('Son paragraf')), {
        sectPr: bothReferences,
        parts: [part('header', p(picture), 1, LOGO_RELS), part('footer', p(r('Dosya No: ') + r('2026/0000', '<w:i/>')))],
        files: LOGO
    });

    assert.deepEqual(topLevel(doc).map(el => el.localName), ['header', 'paragraph', 'paragraph', 'footer']);
    assert.ok(content.startsWith('￼'));
    assert.equal(parsed.header.length, 1);
    const found = pieces(doc, content);
    assert.deepEqual(found.map(({ section, text }) => [section, text]), [
        ['header', '￼'],
        ['body', 'Gövde'], ['body', ' metni'], ['body', 'Son paragraf'],
        ['footer', 'Dosya No: '], ['footer', '2026/0000']
    ]);
    let offset = 0;
    for (const { start, text } of found) {
        assert.equal(start, offset);
        offset += text.length;
    }
    assert.equal(offset, content.length);
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('a footer with a picture and text keeps both, in order', async () => {
    const { doc, content } = await convertBody(p(r('Gövde')), {
        sectPr: reference('footer'),
        parts: [part('footer', p(r('Önce ') + picture + r(FOOTER_TEXT)), 1, LOGO_RELS)],
        files: LOGO
    });
    const footerEl = topLevel(doc).at(-1);
    assert.equal(footerEl.localName, 'footer');
    assert.deepEqual([...footerEl.firstElementChild.children].map(el => el.localName), ['content', 'image', 'content']);
    assert.deepEqual(texts(footerEl, content), ['Önce ', '￼', FOOTER_TEXT]);
    assert.ok(content.endsWith('Önce ￼' + FOOTER_TEXT));
    assert.deepEqual(imagesIn(footerEl), [logo(72, 36)]);
});

for (const [name, run, size] of [['a VML picture (w:pict)', vmlPictureRun(), [120, 40]],
    ['DrawingML with a VML fallback (mc:AlternateContent)', alternatePictureRun(), [72, 36]]]) {
    test(`${name} in a header embeds the picture once`, async () => {
        const { doc, parsed } = await convertBody(p(r('Gövde')), {
            sectPr: reference('header'),
            parts: [part('header', p(run + r(HEADER_TEXT)), 1, LOGO_RELS)],
            files: LOGO
        });
        assert.deepEqual(imagesIn(topLevel(doc)[0]), [logo(...size)]);
        assert.equal(doc.querySelectorAll('image').length, 1);
        assert.deepEqual(findingSummaries(parsed.findings), []);
    });
}

test('a header picture with no relationship or no media part gives missingImage in the header and writes no <image>', async () => {
    for (const [name, options] of [
        ['no relationship', { parts: [part('header', p(r(HEADER_TEXT) + picture))] }],
        ['no media part', { parts: [part('header', p(r(HEADER_TEXT) + picture), 1, LOGO_RELS)] }]
    ]) {
        const { doc, content, parsed } = await convertBody(p(r('Gövde')), { sectPr: reference('header'), ...options });
        assert.equal(doc.querySelectorAll('image').length, 0, name);
        assert.deepEqual(texts(topLevel(doc)[0], content), [HEADER_TEXT], name);
        assert.deepEqual(plain(parsed.findings), [{
            code: 'missingImage', count: 1, examples: [HEADER_TEXT],
            occurrences: [{ example: HEADER_TEXT, location: { kind: 'header' } }]
        }], name);
    }
});

test('a header holding only a missing picture writes no header and gives missingImage', async () => {
    const { xml, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', p(picture))]
    });
    assert.ok(!xml.includes('<header'));
    assert.equal(parsed.header, null);
    assert.deepEqual(plain(parsed.findings).map(({ code, occurrences }) => [code, occurrences.map(o => o.location)]),
        [['missingImage', [{ kind: 'header' }]]]);
});

test('a GIF logo in a header is embedded as-is and gives unsupportedImage; a chart gives unconvertedDrawing', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: bothReferences,
        parts: [part('header', p(pictureRun('rIdGif') + r(HEADER_TEXT)), 1, { rIdGif: 'media/logo.gif' }),
            part('footer', p(r(FOOTER_TEXT) + chartRun))],
        files: { 'word/media/logo.gif': GIF }
    });
    const [headerEl] = topLevel(doc);
    assert.deepEqual(imagesIn(headerEl).map(image => image.data), [GIF.toString('base64')]);
    assert.deepEqual(texts(topLevel(doc).at(-1), content), [FOOTER_TEXT]);
    assert.deepEqual(plain(parsed.findings).map(({ code, examples, occurrences }) => [code, examples, occurrences.map(o => o.location)]), [
        ['unsupportedImage', ['GIF'], [{ kind: 'header' }]],
        ['unconvertedDrawing', [FOOTER_TEXT], [{ kind: 'footer' }]]
    ]);
});

test('a picture in a header field\'s result is left out with the field, and not decoded', async () => {
    const { doc, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', p(r(HEADER_TEXT) + complexField('INCLUDEPICTURE "logo.png"', 'x').replace(r('x'), picture)), 1, LOGO_RELS)],
        files: LOGO
    });
    assert.equal(doc.querySelectorAll('image').length, 0);
    assert.deepEqual(plain(parsed.images), {});
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a watermark in the header is left out and noted, and its picture is not decoded', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header'),
        parts: [part('header', p(watermarkRun()) + p(r(HEADER_TEXT)), 1, LOGO_RELS)],
        files: LOGO
    });
    assert.equal(doc.querySelectorAll('image').length, 0);
    assert.ok(content.startsWith(HEADER_TEXT) || content.startsWith('​' + HEADER_TEXT));
    assert.deepEqual(plain(parsed.images), {});
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a first-page header holding only a picture is not converted, and is noted', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: reference('header', 1) + reference('header', 2, 'first') + '<w:titlePg/>',
        parts: [part('header', p(r(HEADER_TEXT))), part('header', p(picture), 2, LOGO_RELS)],
        files: LOGO
    });
    assert.equal(doc.querySelectorAll('image').length, 0);
    assert.equal(content, HEADER_TEXT + 'Gövde');
    assert.deepEqual(plain(parsed.images), {});
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'headersFooters', count: 1, examples: [] }]);
});

test('a letterhead converts to a well-formed .udf whose header holds the logo first', async () => {
    // Synthetic letterhead: a logo paragraph, the office name and address, a footer with a file reference
    const { doc, content, xml } = await convertBody(p(r('T.C.')) + p(r('İSTANBUL ... MAHKEMESİNE')), {
        sectPr: bothReferences,
        parts: [part('header', p(anchoredPictureRun() + r('Ad Soyad', '<w:b/>') + r(' Hukuk Bürosu')) +
            p(r('Örnek Mah. 1. Sok. No: 0')), 1, LOGO_RELS),
        part('footer', p(r(FOOTER_TEXT)))],
        files: LOGO
    });
    // convert() already checked content.xml is well-formed
    assert.ok(/^<\?xml[^>]*>\s*<template/.test(xml));
    const header = doc.querySelector('elements > header');
    assert.equal(header, topLevel(doc)[0]);
    assert.equal(header.querySelectorAll(':scope > paragraph > image').length, 1);
    assert.ok(content.startsWith('￼Ad Soyad Hukuk Bürosu'));
});
