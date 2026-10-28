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
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
    'xmlns:v="urn:schemas-microsoft-com:vml"'
].join(' ');

const RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const WORDML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const MEDIA_CONTENT_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };

/**
 * Build a minimal .docx in memory. Optional parts get the content-type override and
 * document relationship a Word-saved file gives them.
 * @param {string} bodyXml - Contents of <w:body>
 * @param {Object} [options]
 * @param {Object<string,string>} [options.footnotes] - Footnote id -> inner XML (paragraphs)
 * @param {Object<string,Buffer>} [options.media] - File name under word/media/ -> bytes. Each gets
 *   a document relationship with id `rId_<name without extension>`, e.g. image1.png -> rId_image1
 * @param {string} [options.styles] - Contents of <w:styles> in word/styles.xml
 * @param {string} [options.numbering] - Contents of <w:numbering> in word/numbering.xml
 * @param {string} [options.sectPr] - Contents of the body's final <w:sectPr>
 * @param {string} [options.settings] - Contents of <w:settings> in word/settings.xml
 * @returns {Promise<Buffer>}
 */
async function buildDocx(bodyXml, options = {}) {
    const parts = [];
    const defaults = [];
    const overrides = [];
    const rels = [];
    const wordPart = (name, rootTag, inner) => {
        parts.push([`word/${name}`,
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${rootTag} ${NAMESPACES}>${inner}</w:${rootTag}>`]);
        overrides.push(`<Override PartName="/word/${name}" ContentType="${WORDML_CONTENT_TYPE}${rootTag}+xml"/>`);
        rels.push(`<Relationship Id="rId_${rootTag}" Type="${RELATIONSHIP_TYPE}${rootTag}" Target="${name}"/>`);
    };

    const sectPr = options.sectPr === undefined ? '' : `<w:sectPr>${options.sectPr}</w:sectPr>`;
    parts.push(['word/document.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>${bodyXml}${sectPr}</w:body></w:document>`]);

    if (options.styles !== undefined) wordPart('styles.xml', 'styles', options.styles);
    if (options.numbering !== undefined) wordPart('numbering.xml', 'numbering', options.numbering);
    if (options.settings !== undefined) wordPart('settings.xml', 'settings', options.settings);

    if (options.footnotes) {
        const notes = Object.entries(options.footnotes)
            .map(([id, inner]) => `<w:footnote w:id="${id}">${inner}</w:footnote>`)
            .join('');
        wordPart('footnotes.xml', 'footnotes',
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
            notes);
    }

    if (options.media) {
        for (const [name, data] of Object.entries(options.media)) {
            parts.push([`word/media/${name}`, data]);
            rels.push(`<Relationship Id="rId_${name.replace(/\.[^.]*$/, '')}" ` +
                `Type="${RELATIONSHIP_TYPE}image" Target="media/${name}"/>`);
            const extension = name.split('.').pop().toLowerCase();
            const entry = `<Default Extension="${extension}" ` +
                `ContentType="${MEDIA_CONTENT_TYPES[extension] || 'application/octet-stream'}"/>`;
            if (!defaults.includes(entry)) defaults.push(entry);
        }
    }

    if (rels.length > 0) {
        parts.push(['word/_rels/document.xml.rels',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`]);
    }

    // Package parts first, the way Word orders the archive
    const zip = new JSZip();
    zip.file('[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        defaults.join('') +
        `<Override PartName="/word/document.xml" ContentType="${WORDML_CONTENT_TYPE}document.main+xml"/>` +
        overrides.join('') +
        '</Types>');
    zip.file('_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPE}officeDocument" Target="word/document.xml"/>` +
        '</Relationships>');
    for (const [name, data] of parts) zip.file(name, data);

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
