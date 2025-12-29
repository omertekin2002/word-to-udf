/**
 * DOCX Parser Module
 * Parses Word .docx files and extracts content with formatting
 */

class DocxParser {
    constructor() {
        this.zip = null;
        this.document = null;
        this.relationships = {};
        this.images = {};
    }

    /**
     * Parse a DOCX file
     * @param {File} file - The DOCX file to parse
     * @returns {Promise<Object>} Parsed document structure
     */
    async parse(file) {
        try {
            // Load the DOCX as a ZIP
            this.zip = await JSZip.loadAsync(file);

            // Parse relationships to find media files
            await this.parseRelationships();

            // Extract images
            await this.extractImages();

            // Parse the main document content
            const documentXml = await this.zip.file('word/document.xml').async('string');
            this.document = parseXml(documentXml);

            // Extract document elements
            const elements = await this.extractElements();

            return {
                elements: elements,
                images: this.images
            };
        } catch (error) {
            console.error('Error parsing DOCX:', error);
            throw new Error('Failed to parse DOCX file. Make sure it is a valid Word document.');
        }
    }

    /**
     * Parse document relationships
     */
    async parseRelationships() {
        const relsFile = this.zip.file('word/_rels/document.xml.rels');
        if (!relsFile) return;

        const relsXml = await relsFile.async('string');
        const relsDoc = parseXml(relsXml);
        const relationships = relsDoc.getElementsByTagName('Relationship');

        for (const rel of relationships) {
            const id = rel.getAttribute('Id');
            const target = rel.getAttribute('Target');
            const type = rel.getAttribute('Type');
            this.relationships[id] = { target, type };
        }
    }

    /**
     * Extract embedded images
     */
    async extractImages() {
        const mediaFolder = this.zip.folder('word/media');
        if (!mediaFolder) return;

        const mediaFiles = [];
        this.zip.forEach((relativePath, file) => {
            if (relativePath.startsWith('word/media/')) {
                mediaFiles.push({ path: relativePath, file });
            }
        });

        for (const { path, file } of mediaFiles) {
            const data = await file.async('arraybuffer');
            const base64 = arrayBufferToBase64(data);
            const filename = path.replace('word/media/', '');
            this.images[filename] = base64;
        }
    }

    /**
     * Extract all document elements
     * @returns {Array} Array of document elements
     */
    async extractElements() {
        const body = this.document.getElementsByTagNameNS(WORD_NS, 'body')[0];
        if (!body) return [];

        const elements = [];

        for (const child of body.children) {
            const localName = child.localName;

            if (localName === 'p') {
                elements.push(this.parseParagraph(child));
            } else if (localName === 'tbl') {
                elements.push(this.parseTable(child));
            }
        }

        return elements;
    }

    /**
     * Parse a paragraph element
     * @param {Element} para - The paragraph element
     * @returns {Object} Parsed paragraph
     */
    parseParagraph(para) {
        const paragraph = {
            type: 'paragraph',
            alignment: 'left',
            runs: [],
            numbering: null
        };

        // Get paragraph properties
        const pPr = para.getElementsByTagNameNS(WORD_NS, 'pPr')[0];
        if (pPr) {
            // Alignment
            const jc = pPr.getElementsByTagNameNS(WORD_NS, 'jc')[0];
            if (jc) {
                const val = jc.getAttribute('w:val');
                paragraph.alignment = this.mapAlignment(val);
            }

            // Indentation
            const ind = pPr.getElementsByTagNameNS(WORD_NS, 'ind')[0];
            if (ind) {
                paragraph.leftIndent = this.twipsToPoints(ind.getAttribute('w:left') || '0');
                paragraph.rightIndent = this.twipsToPoints(ind.getAttribute('w:right') || '0');
                paragraph.firstLineIndent = this.twipsToPoints(ind.getAttribute('w:firstLine') || '0');
            }

            // Numbering
            const numPr = pPr.getElementsByTagNameNS(WORD_NS, 'numPr')[0];
            if (numPr) {
                const ilvl = numPr.getElementsByTagNameNS(WORD_NS, 'ilvl')[0];
                const numId = numPr.getElementsByTagNameNS(WORD_NS, 'numId')[0];
                if (ilvl && numId) {
                    paragraph.numbering = {
                        level: parseInt(ilvl.getAttribute('w:val') || '0'),
                        numId: numId.getAttribute('w:val')
                    };
                }
            }
        }

        // Get runs (text with formatting)
        const runs = para.getElementsByTagNameNS(WORD_NS, 'r');
        for (const run of runs) {
            const parsedRun = this.parseRun(run);
            if (parsedRun) {
                paragraph.runs.push(parsedRun);
            }
        }

        return paragraph;
    }

    /**
     * Parse a run element (text with formatting)
     * @param {Element} run - The run element
     * @returns {Object|null} Parsed run
     */
    parseRun(run) {
        // Check for images first
        const drawings = run.getElementsByTagNameNS(WORD_NS, 'drawing');
        if (drawings.length > 0) {
            return this.parseDrawing(drawings[0]);
        }

        // Get text content
        const textElements = run.getElementsByTagNameNS(WORD_NS, 't');
        let text = '';
        for (const t of textElements) {
            text += t.textContent || '';
        }

        // Check for tab
        const tabs = run.getElementsByTagNameNS(WORD_NS, 'tab');
        if (tabs.length > 0) {
            return { type: 'tab' };
        }

        // Check for break
        const breaks = run.getElementsByTagNameNS(WORD_NS, 'br');
        if (breaks.length > 0) {
            const breakType = breaks[0].getAttribute('w:type');
            if (breakType === 'page') {
                return { type: 'pageBreak' };
            }
            return { type: 'break' };
        }

        if (!text) return null;

        // Get run properties
        const formatting = {
            type: 'text',
            text: text,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            fontFamily: 'Times New Roman',
            fontSize: 12
        };

        const rPr = run.getElementsByTagNameNS(WORD_NS, 'rPr')[0];
        if (rPr) {
            // Bold
            const b = rPr.getElementsByTagNameNS(WORD_NS, 'b')[0];
            if (b && b.getAttribute('w:val') !== 'false' && b.getAttribute('w:val') !== '0') {
                formatting.bold = true;
            } else if (b && !b.hasAttribute('w:val')) {
                formatting.bold = true;
            }

            // Italic
            const i = rPr.getElementsByTagNameNS(WORD_NS, 'i')[0];
            if (i && i.getAttribute('w:val') !== 'false' && i.getAttribute('w:val') !== '0') {
                formatting.italic = true;
            } else if (i && !i.hasAttribute('w:val')) {
                formatting.italic = true;
            }

            // Underline
            const u = rPr.getElementsByTagNameNS(WORD_NS, 'u')[0];
            if (u) {
                const val = u.getAttribute('w:val');
                if (val && val !== 'none') {
                    formatting.underline = true;
                }
            }

            // Strikethrough
            const strike = rPr.getElementsByTagNameNS(WORD_NS, 'strike')[0];
            if (strike && strike.getAttribute('w:val') !== 'false') {
                formatting.strike = true;
            }

            // Font family
            const rFonts = rPr.getElementsByTagNameNS(WORD_NS, 'rFonts')[0];
            if (rFonts) {
                const ascii = rFonts.getAttribute('w:ascii');
                if (ascii) formatting.fontFamily = ascii;
            }

            // Font size (in half-points, convert to points)
            const sz = rPr.getElementsByTagNameNS(WORD_NS, 'sz')[0];
            if (sz) {
                const val = sz.getAttribute('w:val');
                if (val) formatting.fontSize = parseInt(val) / 2;
            }

            // Color
            const color = rPr.getElementsByTagNameNS(WORD_NS, 'color')[0];
            if (color) {
                const val = color.getAttribute('w:val');
                if (val && val !== 'auto') {
                    formatting.color = '#' + val;
                }
            }
        }

        return formatting;
    }

    /**
     * Parse a drawing element (embedded image)
     * @param {Element} drawing - The drawing element
     * @returns {Object} Parsed image
     */
    parseDrawing(drawing) {
        const image = {
            type: 'image',
            width: 100,
            height: 100,
            data: null
        };

        // Get extent (size)
        const extents = drawing.getElementsByTagNameNS(WORDML_DRAWING_NS, 'extent');
        if (extents.length > 0) {
            const ext = extents[0];
            // EMUs to points (914400 EMUs = 1 inch = 72 points)
            const cx = parseInt(ext.getAttribute('cx') || '0');
            const cy = parseInt(ext.getAttribute('cy') || '0');
            image.width = Math.round(cx / 914400 * 72);
            image.height = Math.round(cy / 914400 * 72);
        }

        // Get embedded image reference
        const blips = drawing.getElementsByTagNameNS(DRAWING_NS, 'blip');
        if (blips.length > 0) {
            const embed = blips[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
            if (embed && this.relationships[embed]) {
                const target = this.relationships[embed].target;
                const filename = target.replace('media/', '');
                if (this.images[filename]) {
                    image.data = this.images[filename];
                }
            }
        }

        return image;
    }

    /**
     * Parse a table element
     * @param {Element} tbl - The table element
     * @returns {Object} Parsed table
     */
    parseTable(tbl) {
        const table = {
            type: 'table',
            rows: [],
            columnWidths: [],
            border: 'borderCell'
        };

        // Get table grid (column widths)
        const tblGrid = tbl.getElementsByTagNameNS(WORD_NS, 'tblGrid')[0];
        if (tblGrid) {
            const gridCols = tblGrid.getElementsByTagNameNS(WORD_NS, 'gridCol');
            for (const col of gridCols) {
                const w = col.getAttribute('w:w');
                table.columnWidths.push(this.twipsToPoints(w || '0'));
            }
        }

        // Get table properties
        const tblPr = tbl.getElementsByTagNameNS(WORD_NS, 'tblPr')[0];
        if (tblPr) {
            const tblBorders = tblPr.getElementsByTagNameNS(WORD_NS, 'tblBorders')[0];
            if (tblBorders) {
                // Check if borders are set to none
                const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
                let allNone = true;
                for (const borderName of borders) {
                    const border = tblBorders.getElementsByTagNameNS(WORD_NS, borderName)[0];
                    if (border) {
                        const val = border.getAttribute('w:val');
                        if (val && val !== 'none' && val !== 'nil') {
                            allNone = false;
                            break;
                        }
                    }
                }
                if (allNone) {
                    table.border = 'borderNone';
                }
            }
        }

        // Get rows
        const rows = tbl.getElementsByTagNameNS(WORD_NS, 'tr');
        for (const row of rows) {
            table.rows.push(this.parseTableRow(row));
        }

        return table;
    }

    /**
     * Parse a table row
     * @param {Element} tr - The table row element
     * @returns {Object} Parsed row
     */
    parseTableRow(tr) {
        const row = {
            type: 'row',
            cells: []
        };

        const cells = tr.getElementsByTagNameNS(WORD_NS, 'tc');
        for (const cell of cells) {
            row.cells.push(this.parseTableCell(cell));
        }

        return row;
    }

    /**
     * Parse a table cell
     * @param {Element} tc - The table cell element
     * @returns {Object} Parsed cell
     */
    parseTableCell(tc) {
        const cell = {
            type: 'cell',
            paragraphs: [],
            colspan: 1,
            rowspan: 1,
            vAlign: 'top',
            bgColor: null
        };

        // Get cell properties
        const tcPr = tc.getElementsByTagNameNS(WORD_NS, 'tcPr')[0];
        if (tcPr) {
            // Grid span (colspan)
            const gridSpan = tcPr.getElementsByTagNameNS(WORD_NS, 'gridSpan')[0];
            if (gridSpan) {
                cell.colspan = parseInt(gridSpan.getAttribute('w:val') || '1');
            }

            // Vertical merge
            const vMerge = tcPr.getElementsByTagNameNS(WORD_NS, 'vMerge')[0];
            if (vMerge) {
                const val = vMerge.getAttribute('w:val');
                if (!val || val === 'continue') {
                    cell.vMergeContinue = true;
                } else if (val === 'restart') {
                    cell.vMergeStart = true;
                }
            }

            // Vertical alignment
            const vAlign = tcPr.getElementsByTagNameNS(WORD_NS, 'vAlign')[0];
            if (vAlign) {
                cell.vAlign = vAlign.getAttribute('w:val') || 'top';
            }

            // Background color
            const shd = tcPr.getElementsByTagNameNS(WORD_NS, 'shd')[0];
            if (shd) {
                const fill = shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') {
                    cell.bgColor = '#' + fill;
                }
            }
        }

        // Get paragraphs in cell
        const paragraphs = tc.getElementsByTagNameNS(WORD_NS, 'p');
        for (const para of paragraphs) {
            cell.paragraphs.push(this.parseParagraph(para));
        }

        return cell;
    }

    /**
     * Map Word alignment value to UDF alignment
     * @param {string} val - Word alignment value
     * @returns {string} Alignment value
     */
    mapAlignment(val) {
        const alignments = {
            'left': 'left',
            'start': 'left',
            'center': 'center',
            'right': 'right',
            'end': 'right',
            'both': 'justify',
            'distribute': 'justify'
        };
        return alignments[val] || 'left';
    }

    /**
     * Convert twips to points
     * @param {string|number} twips - Value in twips
     * @returns {number} Value in points
     */
    twipsToPoints(twips) {
        return Math.round(parseInt(twips) / 20);
    }
}
