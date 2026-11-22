/**
 * Output vocabulary: the generator writes only the elements and attributes the
 * documented UDF format names. A misspelled or invented name passes every other
 * test and only shows up in the UYAP editor, so this test converts a broad corpus
 * through the real parser and generator and checks every element and attribute
 * name in the output (names only, not values).
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM } = require('jsdom');
const { convert, convertBody } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

const { DOMParser } = new JSDOM('').window;

/**
 * Every element the documented format names, with the attributes it names for it, from
 * saidsurucu/UDF-Toolkit Docs.md at commit 15496404:
 * https://github.com/saidsurucu/UDF-Toolkit/blob/15496404cb38ee377ef087cb9b57649c352ed074/Docs.md
 *
 * `written` lists what the generator is meant to write: each must appear somewhere in the
 * corpus output, so the corpus can't quietly stop exercising a branch. `unused` lists what
 * the documentation names but the generator intentionally doesn't write; it is allowed,
 * but not required. An element marked `notWritten` is allowed but not required either.
 *
 * <content> means two things: the CDATA text pool, a child of <template>, and a run of
 * text in a paragraph. The pool is keyed `template > content`; it takes no attributes.
 *
 * Intentionally unused:
 * - <bgImage> (background image) and <space> (an extra space between <content> elements):
 *   nothing in a Word document is converted to them.
 * - <footer pageNumber-*>: the pageNumber-spec format isn't documented well enough to
 *   produce (see the README's "Headers and footers").
 * - <paragraph family size description> (only in the <header> example), SecListTypeLevel1,
 *   <tab family size>, <row border height>: not written, the editor's defaults apply.
 *
 * Left out of the allowlist on purpose: the hypothetical ("Varsayımsal") <data> section
 * and <field> element, whose structure the documentation itself only guesses at. The XML
 * declaration's version and encoding aren't checked: the parser doesn't expose them as
 * attributes.
 */
const DOCUMENTED = {
    // "Kök Eleman"
    template: { written: ['format_id'] },
    // "İçerik Bölümü": the CDATA text pool
    'template > content': { written: [] },
    // "İçerik (<content> elemanı)": a run of text
    content: {
        written: ['startOffset', 'length', 'family', 'size', 'bold', 'italic', 'underline', 'foreground', 'background']
    },
    // "Özellikler Bölümü"
    properties: { written: [] },
    pageFormat: {
        written: ['mediaSizeName', 'leftMargin', 'rightMargin', 'topMargin', 'bottomMargin',
            'paperOrientation', 'headerFOffset', 'footerFOffset']
    },
    bgImage: {
        notWritten: true,
        written: [],
        // bgImageRigtMargin is spelled this way in the format
        unused: ['bgImageSource', 'bgImageData', 'bgImageBottomMargin', 'bgImageUpMargin',
            'bgImageRigtMargin', 'bgImageLeftMargin']
    },
    // "Elemanlar Bölümü"
    elements: { written: ['resolver'] },
    // "Stiller Bölümü"
    styles: { written: [] },
    style: { written: ['name', 'description', 'family', 'size', 'bold', 'italic', 'foreground', 'FONT_ATTRIBUTE_KEY'] },
    // "Üstbilgi" and "Altbilgi"
    header: { written: [] },
    footer: {
        written: [],
        unused: ['pageNumber-spec', 'pageNumber-fontBold', 'pageNumber-fontItalic', 'pageNumber-fontFace',
            'pageNumber-fontSize', 'pageNumber-color']
    },
    // "Paragraf" and "Liste Özellikleri"
    paragraph: {
        written: ['Alignment', 'LeftIndent', 'RightIndent', 'FirstLineIndent', 'LineSpacing', 'SpaceAbove',
            'SpaceBelow', 'TabSet', 'Numbered', 'NumberType', 'Bulleted', 'BulletType', 'ListId', 'ListLevel'],
        unused: ['family', 'size', 'description', 'SecListTypeLevel1']
    },
    // "Resim"
    image: { written: ['imageData', 'startOffset', 'length', 'width', 'height'] },
    // "Sekme"
    tab: { written: ['startOffset', 'length'], unused: ['family', 'size'] },
    // "Tablo", "Satır" and "Hücre"
    table: { written: ['tableName', 'columnCount', 'columnSpans', 'border'] },
    row: { written: ['rowName', 'rowType'], unused: ['border', 'height'] },
    cell: { written: ['colspan', 'align', 'fillColor', 'border', 'borderSpec', 'borderStyle', 'borderWidth', 'borderColor'] },
    // "Boşluk"
    space: { notWritten: true, written: [], unused: ['startOffset', 'length'] },
    // "Sayfa Sonu"
    'page-break': { written: [] }
};

/**
 * Known exceptions: attributes the generator writes that Docs.md does not name. Kept apart
 * from the documented allowlist so they stay visible, and allowed only on these elements.
 *
 * Every documented <content> formatting attribute (family, size, bold, italic, underline,
 * foreground, background) is a Java Swing StyleConstants key name, and these two are StyleConstants
 * keys too, so the editor plausibly reads them. That is inference, not documentation:
 * they're unverified until checked in the official UYAP editor. They're still written,
 * since dropping superscript would turn every note mark into a full-size digit.
 */
const UNDOCUMENTED_EXCEPTIONS = {
    content: [
        'superscript', // footnote and endnote reference marks and note numbers
        'strikethrough' // w:strike
    ]
};

const allowed = element => new Set([
    ...(DOCUMENTED[element].written || []),
    ...(DOCUMENTED[element].unused || []),
    ...(UNDOCUMENTED_EXCEPTIONS[element] || [])
]);

// Small documents that together reach every generator branch that writes markup
const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const cell = (text, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${p(r(text))}</w:tc>`;
const PNG = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
const level = (ilvl, numFmt, lvlText) =>
    `<w:lvl w:ilvl="${ilvl}"><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/></w:lvl>`;
const numPr = numId => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;
const borders = (tag, val, ...edges) =>
    `<w:${tag}>${edges.map(edge => `<w:${edge} w:val="${val}" w:sz="4" w:space="0" w:color="auto"/>`).join('')}</w:${tag}>`;
const ALL_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
const grid = columns => `<w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>`;
const picture = rId => '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const drawing = picture('rId_image1');
const footnoteRef = id => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const part = (kind, inner, rels) => ({
    name: `${kind}1.xml`, root: kind === 'header' ? 'hdr' : 'ftr', type: kind, id: `rId_${kind}1`, inner, rels
});

const FEATURES = [
    {
        // Tracked changes are converted as accepted: inserted text kept, a deleted paragraph and row left out
        name: 'tracked changes',
        body: p(r('Kalan ') + '<w:ins w:id="1" w:author="A"><w:r><w:t>eklenen</w:t></w:r></w:ins>' +
            '<w:del w:id="2" w:author="A"><w:r><w:delText>silinen</w:delText></w:r></w:del>') +
            '<w:p><w:pPr><w:rPr><w:del w:id="3" w:author="A"/></w:rPr></w:pPr>' +
            '<w:del w:id="4" w:author="A"><w:r><w:delText>Silinen paragraf</w:delText></w:r></w:del></w:p>' +
            '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + p(r('Kalan satır')) + '</w:tc></w:tr>' +
            '<w:tr><w:trPr><w:del w:id="5" w:author="A"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
            '<w:p><w:del w:id="6" w:author="A"><w:r><w:delText>Silinen satır</w:delText></w:r></w:del></w:p></w:tc></w:tr></w:tbl>'
    },
    {
        // An equation is written as ordinary text in its linear form
        name: 'equation',
        body: p(r('Faiz: ') + '<m:oMath><m:f><m:num><m:r><m:t>A</m:t></m:r></m:num>' +
            '<m:den><m:r><m:t>365</m:t></m:r></m:den></m:f></m:oMath>' + r(' TL'))
    },
    {
        name: 'run formatting',
        body: p(r('Bold', '<w:b/>') + r('Italic', '<w:i/>') + r('Underline', '<w:u w:val="single"/>') +
            r('Strike', '<w:strike/>') + r('Colour', '<w:color w:val="FF0000"/>') +
            r('Highlight', '<w:highlight w:val="yellow"/>') +
            r('Font', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="28"/>') +
            '<w:r><w:t>Line</w:t><w:br/><w:t>break</w:t></w:r>', '<w:jc w:val="center"/>')
    },
    {
        name: 'paragraph layout, tabs and TabSet',
        body: p(r('Label') + '<w:r><w:tab/></w:r>' + r('Value'),
            '<w:tabs><w:tab w:val="left" w:pos="1985"/></w:tabs>' +
            '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>' +
            '<w:ind w:left="720" w:right="360" w:firstLine="360"/>') +
            '<w:p/>'
    },
    {
        // numId 3 isn't in numbering.xml, so the generator guesses; numId 4's level is numFmt none
        name: 'bulleted, numbered, undefined and unnumbered lists',
        body: p(r('Bullet'), numPr(1)) + p(r('Number'), numPr(2)) + p(r('Guessed'), numPr(3)) + p(r('None'), numPr(4)),
        options: {
            numbering: `<w:abstractNum w:abstractNumId="0">${level(0, 'bullet', '•')}</w:abstractNum>` +
                `<w:abstractNum w:abstractNumId="1">${level(0, 'decimal', '%1.')}</w:abstractNum>` +
                `<w:abstractNum w:abstractNumId="2">${level(0, 'none', '')}</w:abstractNum>` +
                '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
                '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
                '<w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>'
        }
    },
    {
        name: 'table with merged, shaded and aligned cells',
        body: `<w:tbl><w:tblPr>${borders('tblBorders', 'single', ...ALL_EDGES)}</w:tblPr>${grid(3)}` +
            '<w:tr>' + cell('Spanned', '<w:gridSpan w:val="2"/>') +
            cell('Merged', '<w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>') + '</w:tr>' +
            '<w:tr>' + cell('Centred', '<w:vAlign w:val="center"/>') + cell('Plain') +
            cell('', '<w:vMerge/>') + '</w:tr>' +
            '</w:tbl>' +
            p(r('After the table'))
    },
    {
        // The table draws no edge itself: the first cell draws its top and left (borderSpec 9),
        // and the second cell, sharing no drawn edge with it, has none (border="borderNone")
        name: 'table with partly bordered and unbordered cells',
        body: `<w:tbl><w:tblPr>${borders('tblBorders', 'nil', ...ALL_EDGES)}</w:tblPr>${grid(2)}` +
            `<w:tr>${cell('Top and left', borders('tcBorders', 'single', 'top', 'left'))}${cell('Unbordered')}</w:tr>` +
            '</w:tbl>' + p(r('After the table'))
    },
    {
        // A double red frame and dotted grey grid lines: every cell writes its heaviest line as
        // borderStyle, borderWidth and borderColor
        name: 'table with styled border lines',
        body: '<w:tbl><w:tblPr><w:tblBorders>' +
            ['top', 'left', 'bottom', 'right'].map(edge => `<w:${edge} w:val="double" w:sz="6" w:space="0" w:color="FF0000"/>`).join('') +
            ['insideH', 'insideV'].map(edge => `<w:${edge} w:val="dotted" w:sz="4" w:space="0" w:color="808080"/>`).join('') +
            `</w:tblBorders></w:tblPr>${grid(2)}` +
            `<w:tr>${cell('Frame')}${cell('Lines')}</w:tr></w:tbl>` + p(r('After the table'))
    },
    {
        // No grid (equal column widths), no edges (border="borderNone" on the table), and a cell
        // with no paragraph at all
        name: 'borderless table with no grid and an empty cell',
        body: `<w:tbl><w:tblPr>${borders('tblBorders', 'nil', ...ALL_EDGES)}</w:tblPr>` +
            `<w:tr>${cell('Text')}<w:tc></w:tc></w:tr></w:tbl>` + p(r('After the table'))
    },
    {
        // A <table> inside a <cell>, between the cell's paragraphs ("Hücreler paragraflar ve iç
        // içe tablolar içerebilir"), with its own merged, shaded and partly bordered cells
        name: 'table nested in a table cell',
        body: `<w:tbl><w:tblPr>${borders('tblBorders', 'nil', ...ALL_EDGES)}</w:tblPr>${grid(1)}<w:tr><w:tc>` +
            p(r('Parties')) +
            `<w:tbl><w:tblPr>${borders('tblBorders', 'double', ...ALL_EDGES)}</w:tblPr>${grid(2)}` +
            '<w:tr>' + cell('Claimant', '<w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>') +
            cell('Name', '<w:vAlign w:val="bottom"/>') + '</w:tr>' +
            '<w:tr>' + cell('', '<w:vMerge/>') + cell('Address') + '</w:tr>' +
            '<w:tr>' + cell('Signature', '<w:gridSpan w:val="2"/>') + '</w:tr></w:tbl>' +
            p(r('End of cell')) + '</w:tc></w:tr></w:tbl>' + p(r('After the table'))
    },
    {
        name: 'table cell with a tab, a picture and a footnote mark',
        body: `<w:tbl>${grid(1)}<w:tr><w:tc>` +
            p(r('Label') + '<w:r><w:tab/></w:r>' + drawing + r('Noted') + footnoteRef(1)) +
            '</w:tc></w:tr></w:tbl>' + p(r('After the table')),
        options: { media: { 'image1.png': PNG }, footnotes: { 1: p(r('Cell footnote')) } }
    },
    {
        name: 'inline PNG',
        body: p(r('Before') + drawing + r('After')),
        options: { media: { 'image1.png': PNG } }
    },
    {
        name: 'empty document',
        body: ''
    },
    {
        name: 'footnote, endnote and manual page break',
        body: p(r('Noted') + footnoteRef(1)) +
            p(r('Page one') + '<w:r><w:br w:type="page"/></w:r>' + r('Page two') +
                '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r>'),
        options: {
            footnotes: { 1: p(r('Footnote text')) },
            endnotes: { 1: p(r('Endnote text')) }
        }
    },
    {
        name: 'formatted two-paragraph footnote and endnote',
        body: p(r('Cited') + footnoteRef(1) +
            '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r>'),
        options: {
            footnotes: {
                1: p('<w:r><w:footnoteRef/></w:r>' + r(' ') + r('Court', '<w:b/>') + r(' case', '<w:i/><w:u w:val="single"/>') +
                    '<w:r><w:br/></w:r>' + r('E. 2020/1', '<w:strike/><w:color w:val="1F4E79"/>') + '<w:r><w:tab/></w:r>' +
                    r('K. 2021/2', '<w:highlight w:val="yellow"/><w:rFonts w:ascii="Arial"/><w:sz w:val="18"/>')) +
                    p(r('Second paragraph', '<w:shd w:val="clear" w:fill="FFF2CC"/>')) + p('')
            },
            endnotes: { 1: p('<w:r><w:endnoteRef/></w:r>' + r('Endnote', '<w:b/>')) + p(r('more', '<w:i/>')) }
        }
    },
    {
        // The header's logo is an <image> inside <header>, related through the header part's own rels
        name: 'header and footer',
        body: p(r('Body')),
        options: {
            sectPr: '<w:headerReference w:type="default" r:id="rId_header1"/>' +
                '<w:footerReference w:type="default" r:id="rId_footer1"/>' +
                '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
                '<w:pgMar w:top="1418" w:right="1134" w:bottom="1418" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/>',
            parts: [part('header', p(picture('rIdLogo') + r('Header text')), { rIdLogo: 'media/logo.png' }),
                part('footer', p(r('Footer text')))],
            files: { 'word/media/logo.png': PNG }
        }
    }
];

/**
 * Every element and attribute name in a converted document, as [element, attribute|null] pairs
 * @param {Document} doc
 * @returns {Array<[string, string|null]>}
 */
function names(doc) {
    const found = [];
    for (const el of [doc.documentElement, ...doc.documentElement.getElementsByTagName('*')]) {
        const element = el.tagName === 'content' && el.parentNode === doc.documentElement
            ? 'template > content'
            : el.tagName;
        found.push([element, null]);
        for (const attribute of el.attributes) found.push([element, attribute.name]);
    }
    return found;
}

/** Names in a document that are neither documented nor a known exception */
function offenders(doc) {
    const bad = new Set();
    for (const [element, attribute] of names(doc)) {
        if (!Object.hasOwn(DOCUMENTED, element)) {
            bad.add(`<${element}>`);
        } else if (attribute !== null && !allowed(element).has(attribute)) {
            bad.add(`<${element} ${attribute}>`);
        }
    }
    return [...bad];
}

// Convert the corpus once and share it between the tests
const CORPUS = [
    ...petitions.map(petition => ({ name: petition.name, convert: () => petition.build().then(convert) })),
    ...FEATURES.map(feature => ({ name: feature.name, convert: () => convertBody(feature.body, feature.options) }))
];
let converted;
const corpus = () => {
    converted ??= Promise.all(CORPUS.map(async item => ({ name: item.name, ...(await item.convert()) })));
    return converted;
};

for (const item of CORPUS) {
    test(`${item.name}: every element and attribute written is documented or a known exception`, async () => {
        const { doc } = (await corpus()).find(result => result.name === item.name);
        const bad = offenders(doc);
        assert.deepEqual(bad, [],
            `${item.name}: the generator wrote ${bad.join(', ')}, which the documented UDF format ` +
            "(Docs.md at 15496404) doesn't name. Add it to DOCUMENTED with its Docs.md section, " +
            'or stop writing it.');
    });
}

test('the corpus writes every documented element and attribute the generator is meant to write', async () => {
    const seen = new Set();
    for (const { doc } of await corpus()) {
        for (const [element, attribute] of names(doc)) seen.add(attribute === null ? `<${element}>` : `<${element} ${attribute}>`);
    }

    const expected = [];
    for (const [element, { written, notWritten }] of Object.entries(DOCUMENTED)) {
        if (!notWritten) expected.push(`<${element}>`);
        for (const attribute of written) expected.push(`<${element} ${attribute}>`);
    }
    for (const [element, attributes] of Object.entries(UNDOCUMENTED_EXCEPTIONS)) {
        for (const attribute of attributes) expected.push(`<${element} ${attribute}>`);
    }

    const missing = expected.filter(name => !seen.has(name));
    assert.deepEqual(missing, [],
        `the corpus no longer exercises ${missing.join(', ')}: add a feature body that writes it, ` +
        "or move it to `unused` if the generator shouldn't write it");
});

test('the corpus writes an <image> inside <header>, as Docs.md\'s Üstbilgi section shows', async () => {
    const { doc } = (await corpus()).find(result => result.name === 'header and footer');
    assert.equal(doc.querySelectorAll('elements > header > paragraph > image').length, 1);
});

test('the vocabulary check names an undocumented element or attribute', () => {
    const doc = new DOMParser().parseFromString(
        '<template format_id="1.8"><content bold="true" /><elements resolver="hvl-default"><paragraph Alignment="0" Shading="1">' +
        '<content startOffset="0" length="1" superscript="true" /><note /></paragraph>' +
        '<table superscript="true" /></elements></template>', 'application/xml');
    assert.deepEqual(offenders(doc),
        ['<template > content bold>', '<paragraph Shading>', '<note>', '<table superscript>']);
});
