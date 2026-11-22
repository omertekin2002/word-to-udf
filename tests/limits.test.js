const { test } = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('node:zlib');
const JSZip = require('jszip');
const { loadConverter, buildDocx, declareSize } = require('./helpers');

const MiB = 1024 * 1024;
const DOCUMENT = 'word/document.xml';

// JSZip feeds a part's compressed bytes to the decompressor in blocks of 16 KiB, and one
// DEFLATE block can expand at most 1032-fold, so that is the most one block can produce
// after the reader has asked the stream to stop
const BLOCK = 16 * 1024;
const MAX_BLOCK_OUTPUT = BLOCK * 1032;

const p = inner => `<w:p>${inner}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const picture = rId => '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const textBox = inner => '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const PNG = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
const JPEG = Buffer.from([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46]);

// The prototype of JSZip's entries, whose internalStream() every read goes through
const ZIP_OBJECT = Object.getPrototypeOf(new JSZip().file('probe', '').file('probe'));

/**
 * Record, for each part decompressed during the test, how many times it was, how many
 * bytes its last stream produced and whether that ran to the end. A part missing from the
 * map was never decompressed.
 * @returns {Map<string, {times: number, produced: number, ended: boolean}>}
 */
function watchReads(t) {
    const original = ZIP_OBJECT.internalStream;
    const reads = new Map();
    t.mock.method(ZIP_OBJECT, 'internalStream', function (type) {
        const stream = original.call(this, type);
        const read = { times: (reads.has(this.name) ? reads.get(this.name).times : 0) + 1, produced: 0, ended: false };
        reads.set(this.name, read);
        stream.on('data', chunk => { read.produced += chunk.length; }).on('end', () => { read.ended = true; });
        return stream;
    });
    return reads;
}

/**
 * Make a part loaded during the test declare no uncompressed size, as if JSZip kept none,
 * so only the streaming checks apply. JSZip itself compares a finished part's length with
 * that size, so only a part that is stopped before its end can do without it.
 */
function hideDeclaredSize(t, name) {
    const original = JSZip.loadAsync;
    t.mock.method(JSZip, 'loadAsync', async function (...args) {
        const zip = await original.apply(this, args);
        zip.file(name)._data.uncompressedSize = undefined;
        return zip;
    });
}

/**
 * Give a stream that wasn't stopped time to keep producing, or to reach its end
 */
const afterStreams = () => new Promise(resolve => setTimeout(resolve, 300));

async function assertRejects(parser, input, code) {
    await assert.rejects(parser.parse(input), error => {
        assert.equal(error.code, code, error.message);
        return true;
    });
}

/**
 * A ZIP archive with one DEFLATE entry of `size` zero bytes, built without holding them in
 * memory; it declares `declared` bytes and a CRC of 0, which JSZip doesn't check by default
 */
async function zeroBomb(name, size, declared) {
    const deflate = zlib.createDeflateRaw();
    const compressed = [];
    deflate.on('data', chunk => compressed.push(chunk));
    const done = new Promise((resolve, reject) => deflate.on('end', resolve).on('error', reject));
    const chunk = Buffer.alloc(MiB);
    for (let written = 0; written < size; written += chunk.length) {
        if (!deflate.write(chunk)) await new Promise(resolve => deflate.once('drain', resolve));
    }
    deflate.end();
    await done;
    const data = Buffer.concat(compressed);

    const fileName = Buffer.from(name);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(declared, 22);
    local.writeUInt16LE(fileName.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(declared, 24);
    central.writeUInt16LE(fileName.length, 28);
    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50, 0);
    end.writeUInt16LE(1, 8);
    end.writeUInt16LE(1, 10);
    end.writeUInt32LE(central.length + fileName.length, 12);
    end.writeUInt32LE(local.length + fileName.length + data.length, 16);
    return Buffer.concat([local, fileName, data, central, fileName, end]);
}

// 256 MiB of zeros, past the default per-part limit, in about 256 KiB
const BOMB_SIZE = 256 * MiB;
let bomb;
const bombDeclaring = declared => (bomb || (bomb = zeroBomb(DOCUMENT, BOMB_SIZE, BOMB_SIZE)))
    .then(archive => declareSize(archive, DOCUMENT, declared));

/**
 * Text that compresses poorly, so a stored or deflated part of it is about as long as the text
 */
function noise(length) {
    let seed = 1;
    let text = '';
    while (text.length < length) {
        seed = (seed * 1103515245 + 12345) % 2147483648;
        text += 'abcdefghijklmnopqrstuvwxyz '[seed % 27];
    }
    return text;
}

// The parser logs the underlying error before rethrowing; keep the test output clean
test.beforeEach(t => t.mock.method(console, 'error', () => {}));

test('the limits default to 100 MiB per part and 300 MiB per parse', () => {
    const { DocxParser } = loadConverter();
    assert.deepEqual({ ...new DocxParser().limits }, { part: 100 * MiB, total: 300 * MiB });
});

test('a part declaring more than the per-part limit is refused as tooLarge without being decompressed', async t => {
    const docx = declareSize(await buildDocx(p(r('Dilekçe metni'))), DOCUMENT, 100 * MiB + 1);
    const reads = watchReads(t);
    await assertRejects(new (loadConverter().DocxParser)(), docx, 'tooLarge');
    assert.ok(!reads.has(DOCUMENT), 'word/document.xml was decompressed');
});

test('a part declaring exactly the per-part limit is read', async t => {
    const docx = await buildDocx(p(r('Dilekçe metni')));
    const size = (await JSZip.loadAsync(docx)).file(DOCUMENT)._data.uncompressedSize;
    const parser = new (loadConverter().DocxParser)();
    parser.limits = { part: size, total: 300 * MiB };
    assert.equal((await parser.parse(docx)).elements.length, 1);
    parser.limits = { part: size - 1, total: 300 * MiB };
    await assertRejects(parser, docx, 'tooLarge');
});

test('an oversized optional part, picture or the zip bomb itself is refused, not skipped', async t => {
    const docx = await buildDocx(p(r('Metin')) + p(picture('rId_image1')), {
        styles: '<w:style w:type="paragraph" w:styleId="Normal"/>',
        parts: [{ name: 'comments.xml', root: 'comments', type: 'comments', id: 'rIdComments', inner: '' }],
        media: { 'image1.png': PNG }
    });
    const reads = watchReads(t);
    for (const part of ['word/styles.xml', 'word/comments.xml', 'word/media/image1.png']) {
        reads.clear();
        await assertRejects(new (loadConverter().DocxParser)(), declareSize(docx, part, 0xFFFFFFFE), 'tooLarge');
        assert.ok(!reads.has(part), `${part} was decompressed`);
    }
    await assertRejects(new (loadConverter().DocxParser)(), await bombDeclaring(BOMB_SIZE), 'tooLarge');
});

test('parts that together pass the total limit are refused; a part read twice counts once', async t => {
    const docx = await buildDocx(p(r(noise(3000))), {
        styles: `<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="${noise(3000)}"/></w:style>`,
        settings: '<w:gutterAtTop/>',
        sectPr: '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="567"/>'
    });
    // Every part but [Content_Types].xml is read, settings.xml more than once
    const zip = await JSZip.loadAsync(docx);
    const sizes = Object.values(zip.files).filter(file => !file.dir && file.name !== '[Content_Types].xml')
        .map(file => file._data.uncompressedSize);
    const total = sizes.reduce((sum, size) => sum + size, 0);
    const parser = new (loadConverter().DocxParser)();

    const reads = watchReads(t);
    parser.limits = { part: 100 * MiB, total };
    assert.equal((await parser.parse(docx)).elements.length, 1);
    assert.equal(reads.size, sizes.length, [...reads.keys()].join(', '));
    assert.ok(reads.get('word/settings.xml').times > 1, 'settings.xml is read again');

    parser.limits = { part: Math.max(...sizes), total: total - 1 };
    await assertRejects(parser, docx, 'tooLarge');
});

test('without declared sizes, a part is stopped once it passes the per-part or the total limit', async t => {
    hideDeclaredSize(t, DOCUMENT);
    const docx = await buildDocx(p(r(noise(256 * 1024))), {
        styles: `<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="${noise(64 * 1024)}"/></w:style>`
    });
    const parser = new (loadConverter().DocxParser)();

    const reads = watchReads(t);
    parser.limits = { part: 100 * 1024, total: 300 * MiB };
    await assertRejects(parser, docx, 'tooLarge');
    await afterStreams();
    // buildDocx stores its parts uncompressed, so each chunk is one 16 KiB block
    assert.ok(reads.get(DOCUMENT).produced <= 100 * 1024 + BLOCK, `produced ${reads.get(DOCUMENT).produced}`);
    assert.equal(reads.get(DOCUMENT).ended, false);

    reads.clear();
    // styles.xml (64 KiB) is read before word/document.xml, which then only has 36 KiB left
    parser.limits = { part: 1 * MiB, total: 100 * 1024 };
    await assertRejects(parser, docx, 'tooLarge');
    await afterStreams();
    assert.ok(reads.get(DOCUMENT).produced <= 36 * 1024 + BLOCK, `produced ${reads.get(DOCUMENT).produced}`);
    assert.equal(reads.get(DOCUMENT).ended, false);
});

test('a part that understates its size is refused as damaged, and decompression stops early', async t => {
    const docx = await buildDocx(p(r(noise(256 * 1024))));
    const reads = watchReads(t);
    await assertRejects(new (loadConverter().DocxParser)(), declareSize(docx, DOCUMENT, 1000), 'corrupt');
    await afterStreams();
    const { produced, ended } = reads.get(DOCUMENT);
    assert.ok(produced <= 1000 + BLOCK, `produced ${produced}`);
    assert.equal(ended, false);
});

test('a zip bomb that understates its size stops within one block of output', async t => {
    const archive = await bombDeclaring(1024);
    assert.ok(archive.length < 1 * MiB, `the bomb is ${archive.length} bytes`);
    const reads = watchReads(t);
    const started = Date.now();
    await assertRejects(new (loadConverter().DocxParser)(), archive, 'corrupt');
    assert.ok(Date.now() - started < 5000, 'refused quickly');
    await afterStreams();
    const { produced, ended } = reads.get(DOCUMENT);
    assert.ok(produced <= 1024 + MAX_BLOCK_OUTPUT, `produced ${produced} of ${BOMB_SIZE}`);
    assert.equal(ended, false);
});

test('a zip bomb with no declared size stops at the per-part limit', async t => {
    hideDeclaredSize(t, DOCUMENT);
    const archive = await bombDeclaring(BOMB_SIZE);
    const reads = watchReads(t);
    const parser = new (loadConverter().DocxParser)();
    parser.limits = { part: 1 * MiB, total: 300 * MiB };
    await assertRejects(parser, archive, 'tooLarge');
    await afterStreams();
    const { produced, ended } = reads.get(DOCUMENT);
    assert.ok(produced <= 1 * MiB + MAX_BLOCK_OUTPUT, `produced ${produced} of ${BOMB_SIZE}`);
    assert.equal(ended, false);
});

test('a parser refused for size converts the next document', async () => {
    const parser = new (loadConverter().DocxParser)();
    parser.limits = { part: 100 * MiB, total: 2000 };
    await assertRejects(parser, await buildDocx(p(r(noise(4000)))), 'tooLarge');
    assert.equal((await parser.parse(await buildDocx(p(r('Dilekçe metni'))))).elements.length, 1);
});

test('only the pictures the document shows are decompressed and returned', async t => {
    // image2.png is related but no a:blip shows it, and it claims to be far too large to read
    const docx = declareSize(await buildDocx(p(r('Önce') + picture('rId_image1')) +
        p(r('Kutu') + textBox(p(picture('rId_image3')))), {
        media: { 'image1.png': PNG, 'image2.png': PNG, 'image3.jpg': JPEG }
    }), 'word/media/image2.png', 0xFFFFFFFE);
    const reads = watchReads(t);
    const parsed = await new (loadConverter().DocxParser)().parse(docx);

    assert.ok(!reads.has('word/media/image2.png'), 'the unused picture was decompressed');
    assert.deepEqual(Object.keys(parsed.images).sort(), ['word/media/image1.png', 'word/media/image3.jpg']);
    assert.equal(parsed.images['word/media/image1.png'], PNG.toString('base64'));
    assert.equal(parsed.images['word/media/image3.jpg'], JPEG.toString('base64'));
    assert.equal(parsed.findings.length, 0, JSON.stringify(parsed.findings));
});

test('only the pictures the converted header shows are decompressed, through its own rels', async t => {
    // The default header relates three pictures and shows one; the first-page header shows another.
    // Each picture not shown claims to be far too large to read.
    const header = (name, id, inner, rels) => ({ name, root: 'hdr', type: 'header', id, inner, rels });
    const rels = { rIdLogo: 'media/logo.png', rIdUnused: 'media/unused.png', rIdBox: 'media/box.png' };
    let docx = await buildDocx(p(r('Gövde')), {
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:headerReference w:type="first" r:id="rIdHeader2"/><w:titlePg/>',
        parts: [
            header('header1.xml', 'rIdHeader1', p(r('Büro') + picture('rIdLogo')) + p(textBox(p(picture('rIdBox')))), rels),
            header('header2.xml', 'rIdHeader2', p(picture('rIdFirst')), { rIdFirst: 'media/first.png' })
        ],
        files: { 'word/media/logo.png': PNG, 'word/media/unused.png': PNG, 'word/media/box.png': PNG, 'word/media/first.png': PNG }
    });
    for (const name of ['unused', 'box', 'first']) docx = declareSize(docx, `word/media/${name}.png`, 0xFFFFFFFE);
    const reads = watchReads(t);
    const parsed = await new (loadConverter().DocxParser)().parse(docx);

    assert.deepEqual([...reads.keys()].filter(name => name.startsWith('word/media/')), ['word/media/logo.png']);
    assert.deepEqual(Object.keys(parsed.images), ['word/media/logo.png']);
    assert.equal(parsed.header[0].runs.filter(run => run.type === 'image' && run.data).length, 1);
});

test('a header picture in a deleted, moved-away or hidden run is not decompressed', async t => {
    const header = { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1',
        inner: p(r('Büro') + `<w:del w:id="1" w:author="A">${picture('rIdDeleted')}</w:del>` +
            `<w:moveFrom w:id="2" w:author="A">${picture('rIdMoved')}</w:moveFrom>` +
            picture('rIdHidden').replace('<w:r>', '<w:r><w:rPr><w:vanish/></w:rPr>')),
        rels: { rIdDeleted: 'media/deleted.png', rIdMoved: 'media/moved.png', rIdHidden: 'media/hidden.png' } };
    let docx = await buildDocx(p(r('Gövde')), {
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>',
        parts: [header],
        files: { 'word/media/deleted.png': PNG, 'word/media/moved.png': PNG, 'word/media/hidden.png': PNG }
    });
    for (const name of ['deleted', 'moved', 'hidden']) docx = declareSize(docx, `word/media/${name}.png`, 0xFFFFFFFE);
    const reads = watchReads(t);
    const parsed = await new (loadConverter().DocxParser)().parse(docx);

    assert.deepEqual([...reads.keys()].filter(name => name.startsWith('word/media/')), []);
    assert.deepEqual(Object.keys(parsed.images), []);
    assert.deepEqual([...parsed.header[0].runs].map(run => run.type), ['text']);
});

test('a picture shown twice is decompressed once', async t => {
    const docx = await buildDocx(p(picture('rId_image1')) + p(picture('rId_image1')), { media: { 'image1.png': PNG } });
    const reads = watchReads(t);
    const parsed = await new (loadConverter().DocxParser)().parse(docx);
    assert.deepEqual([...reads.keys()].filter(name => name.startsWith('word/media/')), ['word/media/image1.png']);
    assert.equal(parsed.elements.flatMap(element => element.runs).filter(run => run.type === 'image' && run.data).length, 2);
});
