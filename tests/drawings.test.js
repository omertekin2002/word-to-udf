const { test } = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { loadConverter, buildDocx, convertBody, textOf } = require('./helpers');

const p = inner => `<w:p>${inner}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;

const O_NS = 'xmlns:o="urn:schemas-microsoft-com:office:office"';

/** A VML shape showing an image part, the way older Word versions and compatibility mode write a picture */
const vmlShape = (rId, style = 'width:120pt;height:40pt') =>
    `<v:shape id="_x0000_i1025" type="#_x0000_t75" style="${style}"><v:imagedata r:id="${rId}" ${O_NS} o:title=""/></v:shape>`;
/** The contents of a w:pict picture, without its run */
const pictEl = (rId, style) => `<w:pict>${vmlShape(rId, style)}</w:pict>`;
/** A run holding a VML picture */
const vmlPicture = (rId, style, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${pictEl(rId, style)}</w:r>`;
/** A run holding an OLE object (an embedded Excel sheet, say) shown by its preview picture */
const oleObject = (rId, style = 'width:200pt;height:80pt') => `<w:r><w:object w:dxaOrig="4000" w:dyaOrig="1600">` +
    `${vmlShape(rId, style)}<o:OLEObject ${O_NS} Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" ` +
    'DrawAspect="Content" ObjectID="_1700000000" r:id="rId_ole"/></w:object></w:r>';
/** The contents of a DrawingML picture, without its run */
const drawingEl = rId => '<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing>';
/** A run holding a drawing with the given graphic data and no picture */
const graphic = (uri, inner) => '<w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="3200400"/>' +
    `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const chart = graphic('http://schemas.openxmlformats.org/drawingml/2006/chart',
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId_chart1"/>');
const smartArt = graphic('http://schemas.openxmlformats.org/drawingml/2006/diagram',
    '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId_dm" r:lo="rId_lo" r:qs="rId_qs" r:cs="rId_cs"/>');
const plainShape = graphic('http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    '<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:bodyPr/></wps:wsp>');
const shapeTextBox = text => graphic('http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    `<wps:wsp><wps:txbx><w:txbxContent>${p(r(text))}</w:txbxContent></wps:txbx></wps:wsp>`);
const vmlTextBox = text => `<w:r><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox><w:txbxContent>${p(r(text))}` +
    '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';

const PNG = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
const PNG2 = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x02]);
const EMF = (() => {
    const bytes = Buffer.alloc(88);
    bytes.writeUInt32LE(1, 0);
    bytes.write(' EMF', 40, 'latin1');
    return bytes;
})();

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const images = doc => [...doc.querySelectorAll('elements image')].map(image => ({
    data: image.getAttribute('imageData'),
    width: image.getAttribute('width'),
    height: image.getAttribute('height')
}));

// The prototype of JSZip's entries, whose internalStream() every read goes through
const ZIP_OBJECT = Object.getPrototypeOf(new JSZip().file('probe', '').file('probe'));

/** The names of the parts decompressed during the test */
function watchReads(t) {
    const original = ZIP_OBJECT.internalStream;
    const reads = new Set();
    t.mock.method(ZIP_OBJECT, 'internalStream', function (type) {
        reads.add(this.name);
        return original.call(this, type);
    });
    return reads;
}

test('a VML picture is embedded with its PNG data and its shape\'s size in points', async () => {
    const { doc, xml, content, parsed } = await convertBody(p(r('Logo: ') + vmlPicture('rId_image1')), {
        media: { 'image1.png': PNG }
    });

    assert.deepEqual(images(doc), [{ data: PNG.toString('base64'), width: '120', height: '40' }]);
    assert.match(xml, new RegExp(`<image startOffset="6" length="1" imageData="${PNG.toString('base64')}" width="120" height="40" />`));
    assert.equal(textOf(doc.querySelector('elements image'), content), '￼');
    assert.deepEqual(plain(parsed.findings), []);
});

test('a VML shape\'s width and height convert from each CSS unit, a bare number being pixels', async () => {
    const cases = [
        ['width:1in;height:2.54cm', 72, 72],
        ['width:25.4mm;height:6pc', 72, 72],
        ['height:96px;width:192', 144, 72],
        [' WIDTH : 30pt ; mso-position-horizontal:center; Height:15.4pt', 30, 15],
        ['position:absolute;margin-left:0;width:.5in;height:1e1pt', 36, 10]
    ];
    for (const [style, width, height] of cases) {
        const { doc } = await convertBody(p(vmlPicture('rId_image1', style)), { media: { 'image1.png': PNG } });
        assert.deepEqual(images(doc).map(image => [image.width, image.height]), [[String(width), String(height)]], style);
    }
});

test('a VML picture whose size is missing, unreadable, zero, negative or infinite keeps the 100×100 default', async () => {
    for (const style of ['', 'width:auto;height:12em', 'width:0;height:-5pt', 'width:1e400pt;height:0.2pt', 'width;height:']) {
        const { doc } = await convertBody(p(vmlPicture('rId_image1', style)), { media: { 'image1.png': PNG } });
        assert.deepEqual(images(doc).map(image => [image.width, image.height]), [['100', '100']], style);
    }
    const { doc } = await convertBody(p(vmlPicture('rId_image1', 'width:50pt;height:huge')), { media: { 'image1.png': PNG } });
    assert.deepEqual(images(doc).map(image => [image.width, image.height]), [['50', '100']]);
});

test('a VML picture in a group is sized through the group\'s coordinate system', async () => {
    const shape = (rId, style) => `<v:shape style="${style}"><v:imagedata r:id="${rId}"/></v:shape>`;
    const group = `<w:r><w:pict><v:group style="width:144pt;height:72pt" coordsize="21600,21600">` +
        `${shape('rId_image1', 'position:absolute;width:21600;height:21600')}` +
        // A nested group, with the default 1000×1000 coordinates, filling the left half
        `<v:group style="position:absolute;width:10800;height:21600">${shape('rId_image2', 'width:500;height:250')}</v:group>` +
        '</v:group></w:pict></w:r>';
    const { parsed } = await convertBody(p(group), { media: { 'image1.png': PNG, 'image2.png': PNG2 } });
    const [image] = plain(parsed.elements[0].runs);
    assert.deepEqual([image.width, image.height], [144, 72]);

    // The nested group's picture, on its own: 500/1000 of 10800/21600 of 144pt, 250/1000 of 72pt
    const nested = group.replace(shape('rId_image1', 'position:absolute;width:21600;height:21600'), '');
    const second = plain((await convertBody(p(nested), { media: { 'image2.png': PNG2 } })).parsed.elements[0].runs);
    assert.deepEqual(second.map(run => [run.width, run.height]), [[36, 18]]);

    // A group without a readable size leaves the picture at the default
    const unsized = `<w:r><w:pict><v:group coordsize="100,100">${shape('rId_image1', 'width:50;height:50')}</v:group></w:pict></w:r>`;
    const third = plain((await convertBody(p(unsized), { media: { 'image1.png': PNG } })).parsed.elements[0].runs);
    assert.deepEqual(third.map(run => [run.width, run.height]), [[100, 100]]);
});

test('an OLE object is shown by its preview picture: an EMF preview embeds and gives unsupportedImage', async () => {
    const { doc, parsed } = await convertBody(p(r('Tablo: ') + oleObject('rId_image1')), { media: { 'image1.emf': EMF } });
    assert.deepEqual(images(doc), [{ data: EMF.toString('base64'), width: '200', height: '80' }]);
    assert.deepEqual(plain(parsed.findings), [{ code: 'unsupportedImage', count: 1, examples: ['EMF'] }]);
});

test('an OLE object or VML picture whose preview part is missing gives missingImage and writes no image', async () => {
    const zip = await JSZip.loadAsync(await buildDocx(p(r('Ek tablo') + oleObject('rId_image1')) + p(r('İmza') + vmlPicture('rId_yok')),
        { media: { 'image1.emf': EMF } }));
    zip.remove('word/media/image1.emf');
    const { DocxParser, UdfGenerator } = loadConverter();
    const parsed = await new DocxParser().parse(await zip.generateAsync({ type: 'nodebuffer' }));
    assert.deepEqual(plain(parsed.findings), [{ code: 'missingImage', count: 2, examples: ['Ek tablo', 'İmza'] }]);

    const udf = await JSZip.loadAsync(await (await new UdfGenerator().generate(parsed)).arrayBuffer());
    assert.doesNotMatch(await udf.file('content.xml').async('string'), /<image /);
});

test('an OLE object written with a DrawingML preview is embedded as a drawing', async () => {
    const object = `<w:r><w:object w:dxaOrig="4000" w:dyaOrig="1600">${drawingEl('rId_image1')}` +
        `<o:OLEObject ${O_NS} Type="Embed" ProgID="Equation.3" r:id="rId_ole"/></w:object></w:r>`;
    const { doc, parsed } = await convertBody(p(object), { media: { 'image1.png': PNG } });
    assert.deepEqual(images(doc), [{ data: PNG.toString('base64'), width: '72', height: '36' }]);
    assert.deepEqual(plain(parsed.findings), []);
});

test('alternate content with a DrawingML picture and a VML fallback writes the chosen picture once', async () => {
    const alternate = (requires, choiceId, fallbackId) => `<w:r><mc:AlternateContent>` +
        `<mc:Choice Requires="${requires}">${drawingEl(choiceId)}</mc:Choice>` +
        `<mc:Fallback>${pictEl(fallbackId)}</mc:Fallback></mc:AlternateContent></w:r>`;

    // The same picture, as Word writes it
    let result = await convertBody(p(alternate('wps', 'rId_image1', 'rId_image1')), { media: { 'image1.png': PNG } });
    assert.deepEqual(images(result.doc), [{ data: PNG.toString('base64'), width: '72', height: '36' }]);

    // Different parts tell the branches apart: the supported choice wins, and the fallback isn't read
    result = await convertBody(p(alternate('wps', 'rId_image1', 'rId_image2')), { media: { 'image1.png': PNG, 'image2.png': PNG2 } });
    assert.deepEqual(images(result.doc).map(image => image.data), [PNG.toString('base64')]);
    assert.deepEqual(Object.keys(plain(result.parsed.images)), ['word/media/image1.png']);

    // A choice the parser can't read falls back to the VML picture
    const unsupported = alternate('cx1', 'rId_image1', 'rId_image2')
        .replace('<mc:AlternateContent>', '<mc:AlternateContent xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex">');
    result = await convertBody(p(unsupported), { media: { 'image1.png': PNG, 'image2.png': PNG2 } });
    assert.deepEqual(images(result.doc), [{ data: PNG2.toString('base64'), width: '120', height: '40' }]);
    assert.deepEqual(Object.keys(plain(result.parsed.images)), ['word/media/image2.png']);
    assert.deepEqual(plain(result.parsed.findings), []);
});

test('alternate content inside a drawing or VML picture embeds the branch that is read, not the first picture', async () => {
    const cx1 = 'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"';
    const blip = rId => `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic>`;
    const inDrawing = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic>' +
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        `<mc:AlternateContent ${cx1}><mc:Choice Requires="cx1">${blip('rId_image1')}</mc:Choice>` +
        `<mc:Fallback>${blip('rId_image2')}</mc:Fallback></mc:AlternateContent>` +
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const inPict = `<w:r><w:pict><mc:AlternateContent ${cx1}><mc:Choice Requires="cx1">${vmlShape('rId_image1')}</mc:Choice>` +
        `<mc:Fallback>${vmlShape('rId_image2', 'width:10pt;height:20pt')}</mc:Fallback></mc:AlternateContent></w:pict></w:r>`;
    const { doc, parsed } = await convertBody(p(inDrawing) + p(inPict), { media: { 'image1.png': PNG, 'image2.png': PNG2 } });
    assert.deepEqual(images(doc), [
        { data: PNG2.toString('base64'), width: '72', height: '36' },
        { data: PNG2.toString('base64'), width: '10', height: '20' }
    ]);
    assert.deepEqual(Object.keys(plain(parsed.images)), ['word/media/image2.png']);
    assert.deepEqual(plain(parsed.findings), []);
});

test('a chart, a SmartArt diagram and a shape with no text box give unconvertedDrawing and write no image', async () => {
    const { doc, xml, content, parsed } = await convertBody(
        p(r('Grafik 1: Dava sayıları') + chart) + p(r('Şema') + smartArt) + p(plainShape) + p(r('Sonraki paragraf')));

    assert.doesNotMatch(xml, /<image /);
    assert.equal(content.includes('￼'), false);
    assert.equal(doc.querySelectorAll('elements image').length, 0);
    assert.deepEqual(plain(parsed.findings), [
        { code: 'unconvertedDrawing', count: 3, examples: ['Grafik 1: Dava sayıları', 'Şema'] }
    ]);
});

test('a VML shape with no picture or text box, such as a horizontal line, gives unconvertedDrawing', async () => {
    const line = `<w:r><w:pict><v:rect ${O_NS} style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t"/></w:pict></w:r>`;
    const { xml, parsed } = await convertBody(p(r('Çizgi') + line));
    assert.doesNotMatch(xml, /<image /);
    assert.deepEqual(plain(parsed.findings), [{ code: 'unconvertedDrawing', count: 1, examples: ['Çizgi'] }]);
});

test('an unconverted drawing in a table cell is reported with its cell paragraph, and not in a merged-away cell', async () => {
    const merged = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p(r('Hücre grafiği') + chart)}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${p(r('Birleşik') + chart)}</w:tc></w:tr></w:tbl>`;
    const { parsed } = await convertBody(merged);
    assert.deepEqual(plain(parsed.findings), [{ code: 'unconvertedDrawing', count: 1, examples: ['Hücre grafiği'] }]);
});

test('text-box shapes, DrawingML and VML, keep their paragraphs and give no unconvertedDrawing', async () => {
    const { content, parsed } = await convertBody(p(r('Çapa') + shapeTextBox('Kutu metni')) + p(r('VML çapa') + vmlTextBox('Eski kutu')));
    assert.ok(content.includes('Kutu metni'));
    assert.ok(content.includes('Eski kutu'));
    assert.deepEqual(plain(parsed.findings), []);
});

test('a VML picture inside a VML text box is embedded once, by the text box\'s own paragraph', async () => {
    const box = `<w:r><w:pict><v:shape style="width:200pt;height:100pt"><v:textbox><w:txbxContent>${p(r('Kutu') + vmlPicture('rId_image1'))}` +
        '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';
    const { doc, parsed } = await convertBody(p(r('Çapa') + box), { media: { 'image1.png': PNG } });
    assert.deepEqual(images(doc), [{ data: PNG.toString('base64'), width: '120', height: '40' }]);
    assert.deepEqual(plain(parsed.findings), []);
});

test('a VML picture, OLE object or chart in a hidden run writes nothing and gives only hiddenText', async () => {
    const hidden = '<w:vanish/>';
    const body = p(r('Görünen') + vmlPicture('rId_image1', undefined, hidden)) +
        p(`<w:r><w:rPr>${hidden}</w:rPr><w:object>${vmlShape('rId_image1')}</w:object></w:r>`) +
        p(chart.replace('<w:r>', `<w:r><w:rPr>${hidden}</w:rPr>`));
    const { xml, parsed } = await convertBody(body, { media: { 'image1.png': PNG } });
    assert.doesNotMatch(xml, /<image /);
    assert.deepEqual(plain(parsed.findings), [{ code: 'hiddenText', count: 3, examples: [] }]);
});

test('only the pictures the document shows are decompressed, VML pictures included', async t => {
    // image2.png is related but nothing shows it; image3.png is only in the fallback that isn't read
    const alternate = `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${drawingEl('rId_image4')}</mc:Choice>` +
        `<mc:Fallback>${pictEl('rId_image3')}</mc:Fallback></mc:AlternateContent></w:r>`;
    const docx = await buildDocx(p(vmlPicture('rId_image1')) + p(oleObject('rId_image5')) + p(alternate), {
        media: { 'image1.png': PNG, 'image2.png': PNG, 'image3.png': PNG, 'image4.png': PNG2, 'image5.emf': EMF }
    });
    const reads = watchReads(t);
    const parsed = await new (loadConverter().DocxParser)().parse(docx);

    assert.deepEqual([...reads].filter(name => name.startsWith('word/media/')).sort(),
        ['word/media/image1.png', 'word/media/image4.png', 'word/media/image5.emf']);
    assert.deepEqual(Object.keys(plain(parsed.images)).sort(),
        ['word/media/image1.png', 'word/media/image4.png', 'word/media/image5.emf']);
});

test('the new finding comes after unsupportedImage, and the UDF is the same without the findings', async () => {
    const { DocxParser, UdfGenerator } = loadConverter();
    const docx = await buildDocx(p(r('Resim') + vmlPicture('rId_yok')) + p(oleObject('rId_image1')) +
        p(r('Grafik') + chart) + tbl([p(r('Hücre') + smartArt)]), { media: { 'image1.emf': EMF } });
    const parsed = await new DocxParser().parse(docx);
    assert.deepEqual(plain(parsed.findings), [
        { code: 'missingImage', count: 1, examples: ['Resim'] },
        { code: 'unsupportedImage', count: 1, examples: ['EMF'] },
        { code: 'unconvertedDrawing', count: 2, examples: ['Grafik', 'Hücre'] }
    ]);

    const xmlOf = async blob => (await JSZip.loadAsync(await blob.arrayBuffer())).file('content.xml').async('string');
    const withFindings = await xmlOf(await new UdfGenerator().generate(parsed));
    const without = await xmlOf(await new UdfGenerator().generate({ ...parsed, findings: [] }));
    assert.equal(withFindings, without);
});
