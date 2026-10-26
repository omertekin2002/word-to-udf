const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr}${runs}</w:p>`;
const r = text => `<w:r><w:t>${text}</w:t></w:r>`;
const tbl = rows => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    rows.map(cells => `<w:tr>${cells.map(c => `<w:tc>${c}</w:tc>`).join('')}</w:tr>`).join('') +
    '</w:tbl>';
const sdt = inner => `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const occurrences = (text, needle) => text.split(needle).length - 1;

function contentElements(doc, root = doc) {
    return [...root.querySelectorAll('content[startOffset]')]
        .filter(el => el.closest('elements'));
}

// Text of each top-level UDF paragraph (body paragraphs are not separated in the content)
function paragraphTexts(doc, content) {
    return [...doc.querySelectorAll('elements > paragraph')]
        .map(para => contentElements(doc, para).map(el => textOf(el, content)).join(''));
}

function cellTexts(doc, content) {
    return [...doc.querySelectorAll('elements cell')]
        .map(c => contentElements(doc, c).map(el => textOf(el, content)).join(''));
}

// A floating text box the way Word writes it: a DrawingML shape, with a VML copy as fallback
const textBox = inner => '<w:r><mc:AlternateContent>' +
    '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';

test('text inside a block-level content control is kept, in document order', async () => {
    const { doc, content } = await convertBody(sdt(p(r('Inside SDT'))) + p(r('After')));
    assert.equal(content, 'Inside SDTAfter');
    assert.deepEqual(paragraphTexts(doc, content), ['Inside SDT', 'After']);
});

test('nested content controls and custom XML blocks are walked', async () => {
    const { doc, content } = await convertBody(
        p(r('One')) + sdt(sdt(p(r('Two')))) + '<w:customXml w:element="party">' + p(r('Three')) + '</w:customXml>');
    assert.deepEqual(paragraphTexts(doc, content), ['One', 'Two', 'Three']);
});

test('text inside a content control in a table cell is kept', async () => {
    const { doc, content } = await convertBody(tbl([[sdt(p(r('Cell SDT')))]]));
    assert.deepEqual(cellTexts(doc, content), ['Cell SDT']);
});

test('rows and cells wrapped in content controls are kept', async () => {
    const body = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc>${p(r('A1'))}</w:tc>${sdt(`<w:tc>${p(r('B1'))}</w:tc>`)}</w:tr>` +
        sdt(`<w:tr><w:tc>${p(r('A2'))}</w:tc><w:tc>${p(r('B2'))}</w:tc></w:tr>`) +
        '</w:tbl>';
    const { doc, content } = await convertBody(body);

    assert.equal(doc.querySelectorAll('elements table row').length, 2);
    assert.deepEqual(cellTexts(doc, content), ['A1', 'B1', 'A2', 'B2']);
});

test('a nested table is flattened into its cell and its text appears once', async () => {
    const { doc, content, parsed } = await convertBody(
        tbl([[p(r('Outer')) + tbl([[p(r('Inner'))]]) + p(r('Tail'))]]));

    assert.equal(parsed.elements.length, 1);
    assert.equal(parsed.elements[0].rows.length, 1);
    assert.equal(doc.querySelectorAll('elements table').length, 1);
    assert.equal(doc.querySelectorAll('elements table row').length, 1);
    assert.equal(doc.querySelectorAll('elements cell').length, 1);
    assert.equal(occurrences(content, 'Inner'), 1, JSON.stringify(content));

    const cell = doc.querySelector('elements cell');
    assert.deepEqual(contentElements(doc, cell).map(el => textOf(el, content)), ['Outer', 'Inner', 'Tail']);
});

test('table, grid and cell properties come from the element itself, not a nested table', async () => {
    const inner = '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="none"/></w:tblBorders></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr>${p(r('Inner'))}</w:tc></w:tr></w:tbl>`;
    const outer = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc>${p(r('Outer'))}${inner}${p('')}</w:tc></w:tr></w:tbl>`;
    const { parsed } = await convertBody(outer);

    const table = parsed.elements[0];
    assert.equal(table.border, 'borderCell');
    assert.deepEqual([...table.columnWidths], [200]);
    assert.equal(table.rows[0].cells[0].colspan, 1);
});

test('text-box text appears once, after the anchoring paragraph', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Anchor') + textBox(p(r('BoxText')))));

    assert.equal(occurrences(content, 'BoxText'), 1, JSON.stringify(content));
    assert.deepEqual(paragraphTexts(doc, content), ['Anchor', 'BoxText']);
    const texts = [...parsed.elements].map(el => [...el.runs].filter(run => run.type === 'text').map(run => run.text).join(''));
    assert.deepEqual(texts, ['Anchor', 'BoxText']);
});

test('a VML-only text box is read once', async () => {
    const { doc, content } = await convertBody(p(r('Anchor') +
        `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${p(r('BoxText'))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`));
    assert.equal(occurrences(content, 'BoxText'), 1, JSON.stringify(content));
    assert.deepEqual(paragraphTexts(doc, content), ['Anchor', 'BoxText']);
});

test('a text box in a table cell stays in that cell, once', async () => {
    const { doc, content } = await convertBody(tbl([[p(r('Anchor') + textBox(p(r('BoxText'))))]]));
    assert.deepEqual(cellTexts(doc, content), ['AnchorBoxText']);
    assert.equal(occurrences(content, 'BoxText'), 1);
});

test('a paragraph without pPr does not take a text box paragraph\'s properties', async () => {
    const boxPara = p(r('Centered'),
        '<w:pPr><w:jc w:val="center"/><w:ind w:left="720"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>');
    const { parsed } = await convertBody(p(r('Anchor') + textBox(boxPara)));

    const [anchor, box] = parsed.elements;
    assert.equal(anchor.alignment, 'left');
    assert.equal(anchor.leftIndent, undefined);
    assert.equal(anchor.numbering, null);
    assert.equal(box.alignment, 'center');
});

test('runs in hyperlinks, insertions, inline content controls and other inline wrappers are kept', async () => {
    const { content } = await convertBody(p(
        r('a') +
        `<w:hyperlink r:id="rId9">${r('b')}</w:hyperlink>` +
        `<w:ins w:id="1" w:author="x">${r('c')}</w:ins>` +
        sdt(r('d')) +
        `<w:smartTag w:element="place">${r('e')}</w:smartTag>` +
        `<w:customXml w:element="name">${r('f')}</w:customXml>` +
        `<w:fldSimple w:instr="PAGE">${r('g')}</w:fldSimple>` +
        `<w:moveTo w:id="2" w:author="x">${r('h')}</w:moveTo>` +
        '<w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="0" w:name="x"/>' +
        `<w:dir w:val="ltr">${r('i')}</w:dir>` +
        `<w:bdo w:val="ltr">${r('j')}</w:bdo>`));
    assert.equal(content, 'abcdefghij');
});

test('deleted and moved-away runs produce no text', async () => {
    const { content } = await convertBody(p(
        r('Keep') +
        '<w:del w:id="1" w:author="x"><w:r><w:tab/><w:delText>Gone</w:delText></w:r></w:del>' +
        `<w:moveFrom w:id="2" w:author="x">${r('Moved')}</w:moveFrom>`));
    assert.equal(content, 'Keep');
});

test('runs inside Office Math are kept', async () => {
    const { content } = await convertBody(p(r('x = ') +
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' + r('y') + '</m:oMath>'));
    assert.equal(content, 'x = y');
});

const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId_image1"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

test('a picture inside a text box embeds once', async () => {
    const { doc } = await convertBody(p(r('Anchor') + textBox(p(drawing))), { media: { 'image1.png': png } });
    assert.equal(doc.querySelectorAll('elements image').length, 1);
});

test('an inline image still embeds', async () => {
    const { doc } = await convertBody(p(r('Before') + drawing + r('After')), { media: { 'image1.png': png } });

    const image = doc.querySelector('elements image');
    assert.ok(image, 'expected an <image> element');
    assert.equal(image.getAttribute('imageData'), png.toString('base64'));
    assert.equal(image.getAttribute('width'), '72');
    assert.equal(image.getAttribute('height'), '36');
});
