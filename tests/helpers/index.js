/**
 * Test helpers: load the browser scripts into a node:vm context, build
 * minimal .docx files in memory, and run the full parser -> generator pipeline.
 */

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { JSDOM } = require('jsdom');

const ROOT = path.resolve(__dirname, '..', '..');
const SCRIPTS = ['js/utils.js', 'js/docx-parser.js', 'js/udf-generator.js'];
const { DOMParser } = new JSDOM('').window;

/**
 * Load the site's scripts, unmodified, the way the page's <script> tags do.
 * @returns {{DocxParser: Function, UdfGenerator: Function}}
 */
function loadConverter() {
    const context = vm.createContext({ DOMParser, JSZip, btoa, console });
    for (const file of SCRIPTS) {
        const code = fs.readFileSync(path.join(ROOT, file), 'utf8');
        vm.runInContext(code, context, { filename: file });
    }
    // Top-level class declarations live in the shared script scope, not on the global object
    return vm.runInContext('({ DocxParser, UdfGenerator })', context);
}

const NAMESPACES = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
].join(' ');

/**
 * Build a minimal .docx in memory.
 * @param {string} bodyXml - Contents of <w:body>
 * @param {Object} [options]
 * @param {Object<string,string>} [options.footnotes] - Footnote id -> inner XML (paragraphs)
 * @returns {Promise<Buffer>}
 */
async function buildDocx(bodyXml, options = {}) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>');
    zip.file('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>');
    zip.file('word/document.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${bodyXml}</w:body></w:document>`);

    if (options.footnotes) {
        const notes = Object.entries(options.footnotes)
            .map(([id, inner]) => `<w:footnote w:id="${id}">${inner}</w:footnote>`)
            .join('');
        zip.file('word/footnotes.xml',
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NAMESPACES}>` +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
            `${notes}</w:footnotes>`);
    }

    return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Run a .docx through DocxParser and UdfGenerator and unpack the resulting .udf.
 * @param {Buffer} docx
 * @returns {Promise<{parsed: Object, xml: string, doc: Document, content: string}>}
 */
async function convert(docx) {
    const { DocxParser, UdfGenerator } = loadConverter();
    const parsed = await new DocxParser().parse(docx);
    const blob = await new UdfGenerator().generate(parsed);
    const udf = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await udf.file('content.xml').async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    assert.equal(doc.getElementsByTagName('parsererror').length, 0, 'content.xml must be well-formed');
    const content = doc.querySelector('template > content').textContent;
    return { parsed, xml, doc, content };
}

/**
 * Convert a document body and check every element's offsets stay inside the content.
 * @param {string} bodyXml
 * @param {Object} [options] - See buildDocx
 */
async function convertBody(bodyXml, options) {
    const result = await convert(await buildDocx(bodyXml, options));
    assertOffsetsInBounds(result);
    return result;
}

/**
 * Every startOffset + length must fall within the CDATA content length.
 */
function assertOffsetsInBounds({ doc, content }) {
    for (const el of doc.querySelectorAll('elements [startOffset]')) {
        const start = Number(el.getAttribute('startOffset'));
        const length = Number(el.getAttribute('length'));
        assert.ok(Number.isInteger(start) && start >= 0, `bad startOffset on <${el.localName}>`);
        assert.ok(Number.isInteger(length) && length >= 0, `bad length on <${el.localName}>`);
        assert.ok(start + length <= content.length,
            `<${el.localName}> ${start}+${length} exceeds content length ${content.length}`);
    }
}

/**
 * Text covered by an element's startOffset/length.
 */
function textOf(el, content) {
    const start = Number(el.getAttribute('startOffset'));
    return content.slice(start, start + Number(el.getAttribute('length')));
}

module.exports = { loadConverter, buildDocx, convert, convertBody, assertOffsetsInBounds, textOf };
