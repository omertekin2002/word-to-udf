/**
 * Utility functions for Word to UDF Converter
 */

/**
 * Format file size in human-readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted file size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Trigger file download from a Blob
 * @param {Blob} blob - The file blob
 * @param {string} filename - The filename for download
 */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Escape special XML characters
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeXml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Wrap text in a CDATA section. Every "]]>" is split across two sections so it
 * can't end the section early; the parsed text is unchanged.
 * @param {string} text - Text to wrap
 * @returns {string} CDATA markup
 */
function cdata(text) {
  return '<![CDATA[' + (text || '').replace(/]]>/g, ']]]]><![CDATA[>') + ']]>';
}

/**
 * Convert ArrayBuffer to Base64 string
 * @param {ArrayBuffer} buffer - The array buffer
 * @returns {string} Base64 encoded string
 */
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * Get file extension from filename
 * @param {string} filename - The filename
 * @returns {string} The extension (without dot)
 */
function getFileExtension(filename) {
  return filename.split('.').pop().toLowerCase();
}

/**
 * Output filename for a converted file: a trailing .docx (any case) becomes .udf
 * @param {string} filename - The input filename
 * @returns {string} The .udf filename
 */
function toUdfFilename(filename) {
  const base = filename.replace(/\.docx$/i, '');
  return (base || 'document') + '.udf';
}

/**
 * Generate a unique ID
 * @returns {string} Unique ID
 */
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise} Promise that resolves after delay
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Parse XML string to Document
 * @param {string} xmlString - XML content
 * @returns {Document} Parsed XML document
 */
function parseXml(xmlString) {
  const parser = new DOMParser();
  return parser.parseFromString(xmlString, 'application/xml');
}

/**
 * Get text content from XML element, handling namespaces
 * @param {Element} element - The XML element
 * @param {string} tagName - Tag name to find
 * @param {string} namespace - XML namespace
 * @returns {string|null} Text content or null
 */
function getXmlText(element, tagName, namespace) {
  const el = element.getElementsByTagNameNS(namespace, tagName)[0];
  return el ? el.textContent : null;
}

/**
 * Direct child elements of an element, optionally filtered by namespace and local name.
 * Unlike getElementsByTagNameNS, this never matches deeper descendants.
 * @param {Element} element - The parent element
 * @param {string} [localName] - Local name to match (any if omitted)
 * @param {string} [namespace] - XML namespace (defaults to Word ML)
 * @returns {Array<Element>} Matching children in document order
 */
function childElements(element, localName, namespace = WORD_NS) {
  return Array.from(element.children).filter(child =>
    child.namespaceURI === namespace && (!localName || child.localName === localName));
}

/**
 * First direct child element with the given local name
 * @param {Element} element - The parent element
 * @param {string} localName - Local name to match
 * @param {string} [namespace] - XML namespace (defaults to Word ML)
 * @returns {Element|null} The child or null
 */
function childElement(element, localName, namespace = WORD_NS) {
  return childElements(element, localName, namespace)[0] || null;
}

/**
 * Work out which cells of a parsed table belong to a vertical merge. A w:vMerge
 * continuation cell joins the cell above it that starts in the same grid column
 * (counting gridSpan), if that cell is part of a merge too. The parser and the
 * generator both use this, so they agree on which cells are written empty.
 * @param {Object} table - Parsed table; only its cells' vMergeStart, vMergeContinue and colspan are read
 * @returns {Array<Array<{joinsAbove: boolean, joinsBelow: boolean, first: Object}>>}
 *   Per row and cell: whether it continues the cell above, whether the cell below
 *   continues it, and the merge's first cell (the cell itself outside a merge)
 */
function verticalMerges(table) {
  const merges = [];
  let above = new Map(); // grid column -> index of the cell starting there in the row above

  for (let i = 0; i < table.rows.length; i++) {
    const cells = table.rows[i].cells;
    const row = [];
    const starts = new Map();
    let column = 0;

    for (let j = 0; j < cells.length; j++) {
      const cell = cells[j];
      const merge = { joinsAbove: false, joinsBelow: false, first: cell };

      const k = above.get(column);
      const upper = k === undefined ? null : table.rows[i - 1].cells[k];
      if (cell.vMergeContinue && upper && (upper.vMergeStart || upper.vMergeContinue)) {
        const upperMerge = merges[i - 1][k];
        upperMerge.joinsBelow = true;
        merge.joinsAbove = true;
        merge.first = upperMerge.first;
      }

      row.push(merge);
      starts.set(column, j);
      column += Math.max(1, cell.colspan || 1);
    }

    merges.push(row);
    above = starts;
  }

  return merges;
}

/**
 * Whether the generator writes a parsed paragraph as a paragraph. In the document body it
 * doesn't when the paragraph holds manual page breaks and nothing that becomes text: only
 * the page breaks are written, and the paragraph's own formatting goes nowhere. A picture
 * whose data is missing is skipped, so it is not text. In a table cell a page break is a
 * line break, and every paragraph is written. The parser and the generator both use this,
 * so paragraph findings are recorded for exactly the paragraphs that are written.
 * @param {Object} paragraph - Parsed paragraph; only its runs' type and data are read
 * @param {boolean} [inCell] - Whether the paragraph is in a table cell
 * @returns {boolean}
 */
function paragraphIsWritten(paragraph, inCell = false) {
  if (inCell) return true;
  const becomesText = run => run.type !== 'pageBreak' && (run.type !== 'image' || Boolean(run.data));
  return !paragraph.runs.some(run => run.type === 'pageBreak') || paragraph.runs.some(becomesText);
}

// Word ML namespace
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PICTURE_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WORDML_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const MARKUP_COMPATIBILITY_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
