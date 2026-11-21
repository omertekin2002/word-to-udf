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

// How long nextFrame() waits for an animation frame before going on without one
const NEXT_FRAME_FALLBACK_MS = 100;

/**
 * Wait until the browser has had a chance to paint what was just put on screen.
 * Runs after the next animation frame; falls back to a timer where there are no
 * animation frames (jsdom) or they are paused (a hidden tab), so it never stalls.
 * @returns {Promise<void>}
 */
function nextFrame() {
  return new Promise(resolve => {
    let done = false;
    const finish = () => {
      if (!done) {
        done = true;
        resolve();
      }
    };
    if (typeof requestAnimationFrame === 'function') {
      // Animation frame callbacks run just before the paint, so continue on the task after it
      requestAnimationFrame(() => setTimeout(finish, 0));
      setTimeout(finish, NEXT_FRAME_FALLBACK_MS);
    } else {
      setTimeout(finish, 0);
    }
  });
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
 * Convert a hex color to a Java-style signed ARGB integer with full alpha
 * @param {string} hex - Hex color (e.g. "#FF0000" or "FF0000")
 * @returns {number} Signed 32-bit ARGB; black (-16777216) for an empty value
 */
function hexToRgbInt(hex) {
  if (!hex) return -16777216; // Black

  hex = hex.replace('#', '');
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);

  // Java-style signed RGB (0xFFRRGGBB as signed 32-bit)
  const rgb = (255 << 24) | (r << 16) | (g << 8) | b;
  return rgb | 0; // Convert to signed 32-bit
}

// <cell borderStyle> for the Word border styles (ST_Border w:val) that have a documented
// equivalent. Any other line style (triple, thinThick*, wave, 3-D and art borders...) is
// approximated as solid.
const CELL_BORDER_STYLES = new Map([
  ['single', 'borderStyle-solid'],
  ['thick', 'borderStyle-solid'],
  ['dotted', 'borderStyle-dotted'],
  ['dashed', 'borderStyle-dashed'],
  ['dashSmallGap', 'borderStyle-dashed'],
  ['dotDash', 'borderStyle-dashed'],
  ['dotDotDash', 'borderStyle-dashed'],
  ['dashDotStroked', 'borderStyle-dashed'],
  ['double', 'borderStyle-double']
]);

// How heavy each <cell borderStyle> is, for lines of equal width: double > solid > dashed > dotted
const CELL_BORDER_STYLE_WEIGHTS = new Map([
  ['borderStyle-dotted', 1],
  ['borderStyle-dashed', 2],
  ['borderStyle-solid', 3],
  ['borderStyle-double', 4]
]);

// Word's default table line: single, ½ pt (4 eighths of a point), automatic (black) color
const DEFAULT_BORDER_LINE = Object.freeze({ val: 'single', size: 4, color: 'auto' });

/**
 * The <cell> border attributes of a parsed border line (see DocxParser.readBorders)
 * @param {{val: string, size: number, color: string}} line - Word's w:val, w:sz in eighths
 *   of a point, and w:color ('auto' or RRGGBB)
 * @returns {{borderStyle: string, borderWidth: string, borderColor: number, documented: boolean}}
 *   documented is false when the style has no documented equivalent and was approximated as solid.
 *   borderWidth is in points, with at least one decimal place ("0.5", "1.0", "2.25").
 */
function cellBorderLine(line) {
  const style = CELL_BORDER_STYLES.get(line.val);
  const points = line.size / 8;
  return {
    borderStyle: style || 'borderStyle-solid',
    borderWidth: Number.isInteger(points * 10) ? points.toFixed(1) : String(points),
    borderColor: line.color === 'auto' ? -16777216 : hexToRgbInt(line.color),
    documented: Boolean(style)
  };
}

/**
 * Whether a parsed border line is Word's default line (single, ½ pt, automatic or black), which
 * the generator leaves to the editor's own default line: no <cell> border attributes
 * @param {{val: string, size: number, color: string}} line
 * @returns {boolean}
 */
function isDefaultBorderLine(line) {
  return line.val === DEFAULT_BORDER_LINE.val && line.size === DEFAULT_BORDER_LINE.size &&
    cellBorderLine(line).borderColor === cellBorderLine(DEFAULT_BORDER_LINE).borderColor;
}

/**
 * The heavier of two border lines: the wider one, else the one whose <cell borderStyle> is
 * heavier (double > solid > dashed > dotted), else the first. A simplified form of Word's
 * conflict rule for a shared edge. Either may be null or undefined (no line).
 * @param {Object|null|undefined} a - A parsed border line
 * @param {Object|null|undefined} b - A parsed border line
 * @returns {Object|null} The heavier line, or null when neither draws one
 */
function heavierBorderLine(a, b) {
  if (!a) return b || null;
  if (!b) return a;
  if (a.size !== b.size) return a.size > b.size ? a : b;
  const weight = line => CELL_BORDER_STYLE_WEIGHTS.get(cellBorderLine(line).borderStyle);
  return weight(b) > weight(a) ? b : a;
}

/**
 * The border lines a parsed cell draws in the UDF: its resolved edges (cell.borders), less the
 * edges a vertical merge joins it by
 * @param {Object} cell - Parsed cell; a cell without parsed edges draws none here
 * @param {{joinsAbove: boolean, joinsBelow: boolean}} merge - The cell's vertical merge (see verticalMerges)
 * @returns {{top: Object|null, right: Object|null, bottom: Object|null, left: Object|null}}
 */
function drawnCellEdges(cell, merge) {
  const edges = cell.borders || {};
  return {
    top: merge.joinsAbove ? null : edges.top || null,
    right: edges.right || null,
    bottom: merge.joinsBelow ? null : edges.bottom || null,
    left: edges.left || null
  };
}

/**
 * The one line a cell's border attributes describe, since the format has one style, width and
 * color per cell: the heaviest of the lines it draws (see heavierBorderLine). The parser and the
 * generator both use this, so the tableBorders note matches what is written.
 * @param {Object} cell - Parsed cell
 * @param {{joinsAbove: boolean, joinsBelow: boolean}} merge - The cell's vertical merge
 * @returns {{line: Object|null, approximated: boolean, mixed: boolean}} line: null when the cell
 *   draws no edge. approximated: a drawn line's style has no documented equivalent. mixed: the
 *   drawn lines differ in written style, width or color, so one was chosen for the whole cell.
 */
function cellBorderChoice(cell, merge) {
  const lines = Object.values(drawnCellEdges(cell, merge)).filter(Boolean);
  const written = lines.map(cellBorderLine);
  const key = attrs => `${attrs.borderStyle} ${attrs.borderWidth} ${attrs.borderColor}`;
  return {
    line: lines.reduce(heavierBorderLine, null),
    approximated: written.some(attrs => !attrs.documented),
    mixed: new Set(written.map(key)).size > 1
  };
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
const VML_NS = 'urn:schemas-microsoft-com:vml';
