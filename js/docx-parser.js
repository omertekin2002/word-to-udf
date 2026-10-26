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
        this.footnotes = {};
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

            // Parse footnotes
            await this.parseFootnotes();

            // Parse the main document content
            const documentXml = await this.zip.file('word/document.xml').async('string');
            this.document = parseXml(documentXml);

            // Extract document elements
            const elements = await this.extractElements();

            return {
                elements: elements,
                images: this.images,
                footnotes: this.footnotes
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
     * Parse footnotes from word/footnotes.xml
     */
    async parseFootnotes() {
        const footnotesFile = this.zip.file('word/footnotes.xml');
        if (!footnotesFile) return;

        try {
            const footnotesXml = await footnotesFile.async('string');
            const footnotesDoc = parseXml(footnotesXml);
            const footnotes = footnotesDoc.getElementsByTagNameNS(WORD_NS, 'footnote');

            for (const footnote of footnotes) {
                const id = footnote.getAttribute('w:id');
                // Skip separator and continuationSeparator footnotes (id 0 and -1)
                if (id === '0' || id === '-1') continue;

                // Extract text content from the footnote
                const paragraphs = footnote.getElementsByTagNameNS(WORD_NS, 'p');
                let footnoteText = '';

                for (const para of paragraphs) {
                    const runs = para.getElementsByTagNameNS(WORD_NS, 'r');
                    for (const run of runs) {
                        const textElements = run.getElementsByTagNameNS(WORD_NS, 't');
                        for (const t of textElements) {
                            footnoteText += t.textContent || '';
                        }
                    }
                    footnoteText += ' ';
                }

                this.footnotes[id] = footnoteText.trim();
            }
        } catch (error) {
            console.warn('Could not parse footnotes:', error);
        }
    }

    /**
     * Extract all document elements
     * @returns {Array} Array of document elements
     */
    async extractElements() {
        const body = this.document.getElementsByTagNameNS(WORD_NS, 'body')[0];
        if (!body) return [];

        return this.parseBlocks(body);
    }

    /**
     * Children of a body, cell, row or table, looking through block-level
     * content controls (w:sdt) and custom XML wrappers
     * @param {Element} container - The containing element
     * @returns {Array<Element>} Word ML children in document order
     */
    contentChildren(container) {
        const children = [];

        for (const child of childElements(container)) {
            if (child.localName === 'customXml') {
                children.push(...this.contentChildren(child));
            } else if (child.localName === 'sdt') {
                const sdtContent = childElement(child, 'sdtContent');
                if (sdtContent) children.push(...this.contentChildren(sdtContent));
            } else {
                children.push(child);
            }
        }

        return children;
    }

    /**
     * Parse the block-level content (paragraphs and tables) of a body, cell or text box
     * @param {Element} container - The containing element
     * @returns {Array<Object>} Parsed paragraphs and tables in document order
     */
    parseBlocks(container) {
        const elements = [];

        for (const child of this.contentChildren(container)) {
            const localName = child.localName;

            if (localName === 'p') {
                const runs = this.paragraphRuns(child);
                elements.push(this.parseParagraph(child, runs));
                // Text boxes anchored in the paragraph follow it as their own blocks
                for (const run of runs) {
                    for (const textBox of this.findTextBoxes(run)) {
                        elements.push(...this.parseBlocks(textBox));
                    }
                }
            } else if (localName === 'tbl') {
                elements.push(this.parseTable(child));
            }
        }

        return elements;
    }

    /**
     * Collect the runs of a paragraph, including runs inside inline wrappers
     * (hyperlinks, insertions, content controls, smart tags, fields, ...).
     * Deleted and moved-away text is skipped.
     * @param {Element} container - The paragraph or an inline wrapper
     * @param {Array<Element>} [runs] - Accumulator
     * @returns {Array<Element>} w:r elements in document order
     */
    paragraphRuns(container, runs = []) {
        const skipped = ['pPr', 'rPr', 'sdtPr', 'sdtEndPr', 'customXmlPr', 'smartTagPr', 'del', 'moveFrom'];

        for (const child of container.children) {
            if (child.namespaceURI === MARKUP_COMPATIBILITY_NS) {
                if (child.localName === 'AlternateContent') {
                    const branch = this.alternateContentBranch(child);
                    if (branch) this.paragraphRuns(branch, runs);
                }
            } else if (child.namespaceURI !== WORD_NS) {
                // e.g. Office Math (m:oMath), which can hold w:r
                this.paragraphRuns(child, runs);
            } else if (child.localName === 'r') {
                runs.push(child);
            } else if (!skipped.includes(child.localName)) {
                this.paragraphRuns(child, runs);
            }
        }

        return runs;
    }

    /**
     * Pick the one branch of mc:AlternateContent to read: the first mc:Choice,
     * or mc:Fallback when there is no choice. Reading both would duplicate content.
     * @param {Element} alternateContent - The mc:AlternateContent element
     * @returns {Element|null} The branch to read
     */
    alternateContentBranch(alternateContent) {
        return childElement(alternateContent, 'Choice', MARKUP_COMPATIBILITY_NS)
            || childElement(alternateContent, 'Fallback', MARKUP_COMPATIBILITY_NS);
    }

    /**
     * Find the text box contents (w:txbxContent) anchored in an element, once each.
     * Only one branch of mc:AlternateContent is read, and text boxes nested in
     * another text box are left for that text box's own parse.
     * @param {Element} element - Element to search (usually a run)
     * @param {Array<Element>} [found] - Accumulator
     * @returns {Array<Element>} w:txbxContent elements in document order
     */
    findTextBoxes(element, found = []) {
        for (const child of element.children) {
            if (child.namespaceURI === WORD_NS && child.localName === 'txbxContent') {
                found.push(child);
            } else if (child.namespaceURI === MARKUP_COMPATIBILITY_NS && child.localName === 'AlternateContent') {
                const branch = this.alternateContentBranch(child);
                if (branch) this.findTextBoxes(branch, found);
            } else {
                this.findTextBoxes(child, found);
            }
        }

        return found;
    }

    /**
     * Parse a paragraph element
     * @param {Element} para - The paragraph element
     * @param {Array<Element>} [runs] - The paragraph's runs, if already collected
     * @returns {Object} Parsed paragraph
     */
    parseParagraph(para, runs = this.paragraphRuns(para)) {
        const paragraph = {
            type: 'paragraph',
            alignment: 'left',
            runs: [],
            numbering: null
        };

        // Get paragraph properties
        const pPr = childElement(para, 'pPr');
        if (pPr) {
            // Alignment
            const jc = childElement(pPr, 'jc');
            if (jc) {
                const val = jc.getAttribute('w:val');
                paragraph.alignment = this.mapAlignment(val);
            }

            // Indentation
            const ind = childElement(pPr, 'ind');
            if (ind) {
                paragraph.leftIndent = this.twipsToPoints(ind.getAttribute('w:left') || '0');
                paragraph.rightIndent = this.twipsToPoints(ind.getAttribute('w:right') || '0');
                paragraph.firstLineIndent = this.twipsToPoints(ind.getAttribute('w:firstLine') || '0');
            }

            // Numbering
            const numPr = childElement(pPr, 'numPr');
            if (numPr) {
                const ilvl = childElement(numPr, 'ilvl');
                const numId = childElement(numPr, 'numId');
                if (ilvl && numId) {
                    paragraph.numbering = {
                        level: parseInt(ilvl.getAttribute('w:val') || '0'),
                        numId: numId.getAttribute('w:val')
                    };
                }
            }
        }

        // Get runs (text with formatting)
        for (const run of runs) {
            paragraph.runs.push(...this.parseRun(run));
        }

        return paragraph;
    }

    /**
     * Parse a run element (text with formatting)
     * @param {Element} run - The run element
     * @returns {Array<Object>} Parsed items in document order (text, tab, break, pageBreak, image, footnoteRef)
     */
    parseRun(run) {
        const items = [];
        let formatting = null;

        for (const child of run.children) {
            const localName = child.localName;

            if (localName === 't') {
                const text = child.textContent || '';
                if (!text) continue;
                const last = items[items.length - 1];
                if (last && last.type === 'text') {
                    last.text += text;
                } else {
                    if (!formatting) formatting = this.parseRunFormatting(run);
                    items.push({ ...formatting, text: text });
                }
            } else if (localName === 'tab') {
                items.push({ type: 'tab' });
            } else if (localName === 'br') {
                const breakType = child.getAttribute('w:type');
                items.push({ type: breakType === 'page' ? 'pageBreak' : 'break' });
            } else if (localName === 'cr') {
                items.push({ type: 'break' });
            } else if (localName === 'footnoteReference') {
                const footnoteId = child.getAttribute('w:id');
                // Use 'in' to check key presence - empty string is still valid
                if (footnoteId && footnoteId in this.footnotes) {
                    items.push({
                        type: 'footnoteRef',
                        id: footnoteId,
                        content: this.footnotes[footnoteId] || '[Empty footnote]'
                    });
                }
            } else if (localName !== 'rPr') {
                // w:drawing, or a drawing wrapped in mc:AlternateContent
                const drawing = localName === 'drawing' ? child
                    : child.getElementsByTagNameNS(WORD_NS, 'drawing')[0];
                // A text box shape is not an image; its content is parsed as blocks
                if (drawing && !drawing.getElementsByTagNameNS(WORD_NS, 'txbxContent').length) {
                    items.push(this.parseDrawing(drawing));
                }
            }
        }

        return items;
    }

    /**
     * Parse a run's character formatting
     * @param {Element} run - The run element
     * @returns {Object} Text item template (without text)
     */
    parseRunFormatting(run) {
        // Get run properties
        const formatting = {
            type: 'text',
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            fontFamily: 'Times New Roman',
            fontSize: 12
        };

        const rPr = childElement(run, 'rPr');
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
        const tblGrid = childElement(tbl, 'tblGrid');
        if (tblGrid) {
            const gridCols = childElements(tblGrid, 'gridCol');
            for (const col of gridCols) {
                const w = col.getAttribute('w:w');
                table.columnWidths.push(this.twipsToPoints(w || '0'));
            }
        }

        // Get table properties
        const tblPr = childElement(tbl, 'tblPr');
        if (tblPr) {
            const tblBorders = childElement(tblPr, 'tblBorders');
            if (tblBorders) {
                // Check if borders are set to none
                const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
                let allNone = true;
                for (const borderName of borders) {
                    const border = childElement(tblBorders, borderName);
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

        // Get rows (only this table's own, not those of nested tables)
        for (const row of this.contentChildren(tbl)) {
            if (row.localName === 'tr') {
                table.rows.push(this.parseTableRow(row));
            }
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

        for (const cell of this.contentChildren(tr)) {
            if (cell.localName === 'tc') {
                row.cells.push(this.parseTableCell(cell));
            }
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
        const tcPr = childElement(tc, 'tcPr');
        if (tcPr) {
            // Grid span (colspan)
            const gridSpan = childElement(tcPr, 'gridSpan');
            if (gridSpan) {
                cell.colspan = parseInt(gridSpan.getAttribute('w:val') || '1');
            }

            // Vertical merge
            const vMerge = childElement(tcPr, 'vMerge');
            if (vMerge) {
                const val = vMerge.getAttribute('w:val');
                if (!val || val === 'continue') {
                    cell.vMergeContinue = true;
                } else if (val === 'restart') {
                    cell.vMergeStart = true;
                }
            }

            // Vertical alignment
            const vAlign = childElement(tcPr, 'vAlign');
            if (vAlign) {
                cell.vAlign = vAlign.getAttribute('w:val') || 'top';
            }

            // Background color
            const shd = childElement(tcPr, 'shd');
            if (shd) {
                const fill = shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') {
                    cell.bgColor = '#' + fill;
                }
            }
        }

        // Get paragraphs in cell. Nested tables can't be represented yet, so
        // their paragraphs are flattened into this cell in reading order.
        cell.paragraphs = this.flattenToParagraphs(this.parseBlocks(tc));

        return cell;
    }

    /**
     * Flatten parsed blocks to paragraphs, reading tables row by row, cell by cell
     * @param {Array<Object>} blocks - Parsed paragraphs and tables
     * @returns {Array<Object>} Paragraphs in reading order
     */
    flattenToParagraphs(blocks) {
        const paragraphs = [];

        for (const block of blocks) {
            if (block.type === 'paragraph') {
                paragraphs.push(block);
            } else if (block.type === 'table') {
                for (const row of block.rows) {
                    for (const cell of row.cells) {
                        // Cell paragraphs are already flattened
                        paragraphs.push(...cell.paragraphs);
                    }
                }
            }
        }

        return paragraphs;
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
