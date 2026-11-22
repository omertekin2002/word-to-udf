/**
 * UDF Generator Module
 * Generates UYAP UDF format from parsed document structure
 */

// <pageFormat> values for anything the document's page setup doesn't give, as written
// before page setup was read from the document
const DEFAULT_PAGE_FORMAT = {
    leftMargin: '42.51968479156494',
    rightMargin: '28.34645652770996',
    topMargin: '14.17322826385498',
    bottomMargin: '14.17322826385498',
    paperOrientation: '1',
    headerFOffset: '20.0',
    footerFOffset: '20.0'
};

// <cell align> for Word's w:vAlign. Top is the default, so it is left out.
const VERTICAL_ALIGN = new Map([
    ['center', 'vcenter'],
    ['bottom', 'bottom']
]);

// <cell borderSpec> edge bits
const BORDER_TOP = 1;
const BORDER_RIGHT = 2;
const BORDER_BOTTOM = 4;
const BORDER_LEFT = 8;
const BORDER_ALL = 15; // top, right, bottom and left

// Line that sets the footnote and endnote sections off from the text above them
const NOTE_SEPARATOR = '────────────────────';

// Font size of the note numbers and of note text without formatting, in points
const NOTE_TEXT_SIZE = 10;

// Roman numeral digits, largest first
const ROMAN_NUMERALS = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
    [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']
];

class UdfGenerator {
    constructor() {
        this.content = '';
        this.pageSetup = {};
        this.elements = [];
        this.currentOffset = 0;
        this.collectedFootnotes = []; // Store footnotes to append at end
        this.footnoteCounter = 0; // Track display numbers for footnotes
        this.collectedEndnotes = []; // Endnotes, appended after everything else
        this.endnoteCounter = 0; // Endnotes are numbered on their own
        this.endnoteFormat = 'lowerRoman';
        this.pendingPageBreak = false; // A section break or page break before waiting for what follows it
        this.bodyStart = 0; // How many elements come before the body's (the header)
    }

    /**
     * Generate UDF file from parsed document
     * @param {Object} document - Parsed document from DocxParser
     * @returns {Promise<Blob>} UDF file as Blob
     */
    async generate(document) {
        // Reset state
        this.content = '';
        this.elements = [];
        this.currentOffset = 0;
        this.collectedFootnotes = [];
        this.footnoteCounter = 0;
        this.collectedEndnotes = [];
        this.endnoteCounter = 0;
        this.endnoteFormat = document.endnoteFormat || 'lowerRoman';
        this.pageSetup = document.pageSetup || {};

        // Header text and pictures come first in the content, before the body (UDF: <header> is the first element)
        if (document.header) {
            this.elements.push(this.headerFooterElement('header', document.header));
        }
        const headerLength = this.content.length;

        // Process all document elements. A block-level page break (a section break, or a paragraph's
        // page break before) is written when the next paragraph or table is: one that ends the
        // document writes none.
        this.pendingPageBreak = false;
        this.bodyStart = this.elements.length;
        for (const element of document.elements) {
            if (element.type === 'pageBreak') {
                this.pendingPageBreak = true;
                continue;
            }
            if (element.type === 'paragraph' || element.type === 'table') this.writePendingPageBreak();
            if (element.type === 'paragraph') {
                this.processParagraph(element);
            } else if (element.type === 'table') {
                this.processTable(element);
            }
        }
        this.pendingPageBreak = false;

        // Append footnotes section at the end if there are any
        if (this.collectedFootnotes.length > 0) {
            this.appendNotesSection(this.collectedFootnotes);
        }

        // Endnotes go after everything else, where Word puts them
        if (this.collectedEndnotes.length > 0) {
            this.appendNotesSection(this.collectedEndnotes);
        }

        // Ensure the body has at least one paragraph
        if (this.content.length === headerLength) {
            this.content += '\u200B'; // Zero-width space
            this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" /></paragraph>`);
            this.currentOffset += 1;
        }

        // Footer text and pictures come last, after the body and the notes (UDF: <footer> is the last element)
        if (document.footer) {
            this.elements.push(this.headerFooterElement('footer', document.footer));
        }

        // Generate final XML
        const xml = this.generateXml();

        // Create ZIP file with content.xml
        const zip = new JSZip();
        zip.file('content.xml', xml);

        return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
    }

    /**
     * Process a paragraph element
     * @param {Object} paragraph - Parsed paragraph
     */
    processParagraph(paragraph) {
        // Page breaks and nothing that becomes text: only the breaks are written
        if (!paragraphIsWritten(paragraph)) {
            for (const run of paragraph.runs) {
                if (run.type === 'pageBreak') this.writePageBreak();
            }
            return;
        }

        const paraContent = [];
        const paraElements = [];
        let paraOffset = this.currentOffset;
        let afterPageBreak = false;
        let paraProps = paragraph;

        // Process runs
        for (const run of paragraph.runs) {
            if (run.type === 'text') {
                const text = run.text;
                const attrs = this.buildContentAttrs(run);
                paraElements.push(`<content startOffset="${this.currentOffset}" length="${text.length}" ${attrs} />`);
                paraContent.push(text);
                this.currentOffset += text.length;
            } else if (run.type === 'tab') {
                paraElements.push(`<tab startOffset="${this.currentOffset}" length="1" />`);
                paraContent.push('\t');
                this.currentOffset += 1;
            } else if (run.type === 'image') {
                if (run.data) {
                    paraElements.push(`<image startOffset="${this.currentOffset}" length="1" imageData="${run.data}" width="${run.width}" height="${run.height}" />`);
                    paraContent.push('\uFFFC'); // Object replacement character
                    this.currentOffset += 1;
                }
            } else if (run.type === 'break') {
                // Reference the newline so the element tree keeps the break
                paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
                paraContent.push('\n');
                this.currentOffset += 1;
            } else if (run.type === 'pageBreak') {
                // Write the runs before the break as a paragraph of their own
                this.content += paraContent.join('');
                const emittedFragment = paraElements.length > 0;
                if (emittedFragment) {
                    this.elements.push(this.buildParagraphElement(paraProps, paraElements));
                }

                this.writePageBreak();

                // Runs after the break continue in a new paragraph on the next page.
                // It is the same Word paragraph, so it must not start another list
                // item or repeat the first-line indent or the space above.
                paraContent.length = 0;
                paraElements.length = 0;
                afterPageBreak = true;
                if (emittedFragment) {
                    paraProps = { ...paragraph, numbering: null, firstLineIndent: 0, spaceBefore: 0 };
                }
            } else if (run.type === 'footnoteRef' || run.type === 'endnoteRef') {
                // Only the superscript mark goes in the text; the note is appended later
                const mark = this.collectNote(run);
                paraElements.push(`<content startOffset="${this.currentOffset}" length="${mark.length}" family="Times New Roman" size="10" superscript="true" />`);
                paraContent.push(mark);
                this.currentOffset += mark.length;
            }
        }

        // If paragraph is empty, add placeholder (unless it ends with a page break)
        if (paraContent.length === 0) {
            if (afterPageBreak) return;
            paraContent.push('\u200B'); // Zero-width space
            paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
            this.currentOffset += 1;
        }

        this.content += paraContent.join('');
        this.elements.push(this.buildParagraphElement(paraProps, paraElements));
    }

    /**
     * Write header or footer paragraphs, with their text and pictures added to the content at
     * the current offset, as body paragraphs are: a picture is an <image> covering one U+FFFC
     * @param {'header'|'footer'} tag - The element to write
     * @param {Array<Object>} paragraphs - Parsed paragraphs, holding only text, tabs, line breaks and pictures
     * @returns {string} The <header> or <footer> element
     */
    headerFooterElement(tag, paragraphs) {
        const bodyElements = this.elements;
        this.elements = [];
        try {
            for (const paragraph of paragraphs) this.processParagraph(paragraph);
            return `<${tag}>${this.elements.join('')}</${tag}>`;
        } finally {
            this.elements = bodyElements;
        }
    }

    /**
     * Write the block-level page break waiting for the paragraph or table about to be written (see
     * generate), as a manual page break is. None before the body's first element, which already
     * starts the first page, nor right after a page break, which would add a blank page.
     */
    writePendingPageBreak() {
        if (!this.pendingPageBreak) return;
        this.pendingPageBreak = false;
        if (this.elements.length === this.bodyStart || this.elements[this.elements.length - 1] === '<page-break />') return;
        this.writePageBreak();
    }

    /**
     * Write a manual page break, after the footnotes collected on the page it ends
     */
    writePageBreak() {
        if (this.collectedFootnotes.length > 0) {
            this.appendNotesSection(this.collectedFootnotes);
            this.collectedFootnotes = []; // Clear for next page
        }

        this.elements.push('<page-break />');
    }

    /**
     * Number a footnote or endnote reference and keep its text and paragraphs for the notes section.
     * Footnotes and endnotes are numbered separately, in reference order (w:id can be
     * non-contiguous): footnotes 1, 2, 3, endnotes in the document's endnote format.
     * @param {Object} run - A footnoteRef or endnoteRef item
     * @returns {string} The reference mark
     */
    collectNote(run) {
        let mark;
        if (run.type === 'endnoteRef') {
            mark = this.formatNoteNumber(++this.endnoteCounter, this.endnoteFormat);
            this.collectedEndnotes.push({ number: mark, text: run.content, paragraphs: run.paragraphs });
        } else {
            mark = String(++this.footnoteCounter);
            this.collectedFootnotes.push({ number: mark, text: run.content, paragraphs: run.paragraphs });
        }
        return mark;
    }

    /**
     * Write a note number in a Word number format
     * @param {number} number - A positive integer
     * @param {string} format - decimal, lowerRoman, upperRoman, lowerLetter or upperLetter; anything else is decimal
     * @returns {string}
     */
    formatNoteNumber(number, format) {
        if (format === 'lowerRoman' || format === 'upperRoman') {
            let roman = '';
            let rest = number;
            for (const [value, digits] of ROMAN_NUMERALS) {
                while (rest >= value) {
                    roman += digits;
                    rest -= value;
                }
            }
            return format === 'upperRoman' ? roman.toUpperCase() : roman;
        }
        if (format === 'lowerLetter' || format === 'upperLetter') {
            // Word's letters run a-z, then aa-zz, aaa-zzz: the letter repeats
            const letter = String.fromCharCode(97 + (number - 1) % 26).repeat(Math.ceil(number / 26));
            return format === 'upperLetter' ? letter.toUpperCase() : letter;
        }
        return String(number);
    }

    /**
     * Append a notes section: an empty line, the separator line, then each note's paragraphs.
     * The first starts with the superscript "<mark>. ". Text keeps its run formatting, tabs
     * become <tab> and line breaks the body's newline; anything else a parsed item holds is
     * skipped. A note without parsed paragraphs (an empty one's placeholder) is its plain text.
     * Note paragraphs are left-aligned, without indents.
     * @param {Array<{number: string, text: string, paragraphs: ?Array<Array<Object>>}>} notes - Footnotes
     *   or endnotes in reference order
     */
    appendNotesSection(notes) {
        // Empty line before separator (element references this newline)
        this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" /></paragraph>`);
        this.content += '\n';
        this.currentOffset += 1;

        // Separator line
        this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="${NOTE_SEPARATOR.length}" family="Times New Roman" size="${NOTE_TEXT_SIZE}" /></paragraph>`);
        this.content += NOTE_SEPARATOR;
        this.currentOffset += NOTE_SEPARATOR.length;

        for (const note of notes) {
            const paragraphs = note.paragraphs && note.paragraphs.some(items => items.length > 0)
                ? note.paragraphs
                : [[{ type: 'text', text: note.text, fontSize: NOTE_TEXT_SIZE }]];

            paragraphs.forEach((items, index) => {
                const paraElements = [];
                const write = (text, element) => {
                    paraElements.push(element);
                    this.content += text;
                    this.currentOffset += text.length;
                };

                // Superscript number
                if (index === 0) {
                    const numPart = note.number + '. ';
                    write(numPart, `<content startOffset="${this.currentOffset}" length="${numPart.length}" family="Times New Roman" size="${NOTE_TEXT_SIZE}" superscript="true" />`);
                }

                for (const item of items) {
                    if (item.type === 'text' && item.text) {
                        write(item.text, `<content startOffset="${this.currentOffset}" length="${item.text.length}" ${this.buildContentAttrs(item)} />`);
                    } else if (item.type === 'tab') {
                        write('\t', `<tab startOffset="${this.currentOffset}" length="1" />`);
                    } else if (item.type === 'break') {
                        write('\n', `<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
                    }
                }

                // A further paragraph with nothing in it, as an empty body paragraph is written
                if (paraElements.length === 0) {
                    write('\u200B', `<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="${NOTE_TEXT_SIZE}" />`);
                }

                this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0">${paraElements.join('')}</paragraph>`);
            });
        }
    }

    /**
     * Build paragraph element XML
     * @param {Object} paragraph - Paragraph data
     * @param {Array} childElements - Child element strings
     * @returns {string} Paragraph XML
     */
    buildParagraphElement(paragraph, childElements) {
        const alignmentMap = {
            'left': '0',
            'center': '1',
            'right': '2',
            'justify': '3'
        };
        const alignment = alignmentMap[paragraph.alignment] || '0';

        let attrs = `Alignment="${alignment}"`;

        // Add indentation
        if (paragraph.leftIndent) {
            attrs += ` LeftIndent="${paragraph.leftIndent}.0"`;
        } else {
            attrs += ` LeftIndent="0.0"`;
        }

        if (paragraph.rightIndent) {
            attrs += ` RightIndent="${paragraph.rightIndent}.0"`;
        } else {
            attrs += ` RightIndent="0.0"`;
        }

        if (paragraph.firstLineIndent) {
            attrs += ` FirstLineIndent="${paragraph.firstLineIndent}.0"`;
        }

        // Spacing: space above and below in points, line spacing as the fraction over single spacing
        const decimal = value => (Number.isInteger(value) ? value.toFixed(1) : String(value));
        if (paragraph.spaceBefore > 0) {
            attrs += ` SpaceAbove="${decimal(paragraph.spaceBefore)}"`;
        }
        if (paragraph.spaceAfter > 0) {
            attrs += ` SpaceBelow="${decimal(paragraph.spaceAfter)}"`;
        }
        const extraLineSpacing = Math.round(((paragraph.lineSpacing || 1) - 1) * 100) / 100;
        if (extraLineSpacing > 0) {
            attrs += ` LineSpacing="${decimal(extraLineSpacing)}"`;
        }

        // Tab stops as position:alignment:leader, positions in points from the left margin
        if (paragraph.tabStops && paragraph.tabStops.length > 0) {
            const stops = paragraph.tabStops.map(stop => `${stop.position.toFixed(1)}:${stop.alignment}:${stop.leader}`);
            attrs += ` TabSet="${stops.join(',')}"`;
        }

        // Handle numbering/bullets
        if (paragraph.numbering) {
            attrs += this.buildListAttrs(paragraph.numbering);
        }

        return `<paragraph ${attrs}>${childElements.join('')}</paragraph>`;
    }

    /**
     * Build the list attributes of a list paragraph
     * @param {Object} numbering - Paragraph numbering (level, numId and, if numbering.xml defines it, format and lvlText)
     * @returns {string} Attributes string, with a leading space, or '' if the level isn't a list
     */
    buildListAttrs(numbering) {
        const { level, numId, format } = numbering;
        const list = ` ListLevel="${level}" ListId="${escapeXml(numId)}"`;

        if (format === undefined) {
            // No list definition in numbering.xml: guess from the numId as before
            return parseInt(numId) % 2 === 0
                ? ` Bulleted="true" BulletType="BULLET_TYPE_ELLIPSE"${list}`
                : ` Numbered="true" NumberType="NUMBER_TYPE_NUMBER_TRE"${list}`;
        }

        if (format === 'none') return '';
        if (format === 'bullet') return ` Bulleted="true" BulletType="BULLET_TYPE_ELLIPSE"${list}`;
        return ` Numbered="true" NumberType="${this.numberType(numbering)}"${list}`;
    }

    /**
     * UDF NumberType for a numbered list level
     * @param {Object} numbering - Paragraph numbering with format and lvlText
     * @returns {string} NUMBER_TYPE_* value
     */
    numberType({ level, format, lvlText }) {
        const families = {
            'decimal': 'NUMBER',
            'lowerLetter': 'CHAR_SMALL',
            'upperLetter': 'CHAR_BIG',
            'lowerRoman': 'ROMAN_SMALL',
            'upperRoman': 'ROMAN_BIG'
        };
        const family = families[format];
        if (!family) return 'NUMBER_TYPE_NUMBER_DOT';

        // The character after this level's placeholder, e.g. ")" in "%2)"
        const placeholder = `%${level + 1}`;
        const at = lvlText.lastIndexOf(placeholder);
        const suffix = at >= 0 ? lvlText.charAt(at + placeholder.length) : '';

        if (suffix === ')') return `NUMBER_TYPE_${family}_PARENTHESIS`;
        if (suffix === '-' && family === 'NUMBER') return 'NUMBER_TYPE_NUMBER_TRE';
        return `NUMBER_TYPE_${family}_DOT`;
    }

    /**
     * Build content element attributes
     * @param {Object} run - Run data
     * @returns {string} Attributes string
     */
    buildContentAttrs(run) {
        const attrs = [];

        attrs.push(`family="${escapeXml(run.fontFamily || 'Times New Roman')}"`);
        attrs.push(`size="${run.fontSize || 12}"`);

        if (run.bold) attrs.push('bold="true"');
        if (run.italic) attrs.push('italic="true"');
        if (run.underline) attrs.push('underline="true"');
        if (run.strike) attrs.push('strikethrough="true"');
        // Superscript (w:vertAlign), as note marks are written; the size stays the run's own
        if (run.superscript) attrs.push('superscript="true"');

        if (run.color) {
            // Convert hex color to RGB integer
            const rgb = this.hexToRgbInt(run.color);
            attrs.push(`foreground="${rgb}"`);
        }

        // Highlighting or run shading, the documented content background color
        if (run.background) attrs.push(`background="${this.hexToRgbInt(run.background)}"`);

        return attrs.join(' ');
    }

    /**
     * Process a table of the document itself
     * @param {Object} table - Parsed table
     */
    processTable(table) {
        this.elements.push(this.tableElement(table));
    }

    /**
     * The <table> element of a parsed table, its cells' text added to the content in reading
     * order. A table nested in a cell is written the same way, in its place in the cell.
     * @param {Object} table - Parsed table
     * @returns {string} The element
     */
    tableElement(table) {
        const columnCount = table.columnWidths.length ||
            (table.rows[0] ? table.rows[0].cells.length : 1);

        // Calculate column spans as proportional values
        let columnSpans;
        if (table.columnWidths.length > 0) {
            const total = table.columnWidths.reduce((a, b) => a + b, 0);
            columnSpans = table.columnWidths.map(w => Math.round((w / total) * 300)).join(',');
        } else {
            // Equal widths
            const equalWidth = Math.round(300 / columnCount);
            columnSpans = Array(columnCount).fill(equalWidth).join(',');
        }

        const merges = verticalMerges(table);
        // The edges each cell draws; the cells of a vertical merge are joined by dropping the edges between them
        const specs = table.rows.map((row, i) => row.cells.map((cell, j) => {
            const merge = merges[i][j];
            return this.cellEdges(cell, table)
                & ~(merge.joinsAbove ? BORDER_TOP : 0)
                & ~(merge.joinsBelow ? BORDER_BOTTOM : 0);
        }));
        const border = specs.some(row => row.some(spec => spec !== 0)) ? 'borderCell' : 'borderNone';
        const rowElements = [];

        for (let i = 0; i < table.rows.length; i++) {
            const row = table.rows[i];
            const cellElements = [];

            for (let j = 0; j < row.cells.length; j++) {
                const cell = row.cells[j];
                const merge = merges[i][j];

                // A continued vertical merge keeps its grid column with an empty cell,
                // since the format has no rowspan. It adds nothing to the content.
                const cellContent = merge.joinsAbove
                    ? this.emptyCellContent()
                    : this.processTableCell(cell);

                let cellAttrs = '';
                if (cell.colspan > 1) {
                    cellAttrs += ` colspan="${cell.colspan}"`;
                }
                // A merged area is shaded the way its first cell is
                const fill = merge.first.bgColor;
                if (fill) {
                    cellAttrs += ` fillColor="${this.hexToRgbInt(fill)}"`;
                }
                const align = VERTICAL_ALIGN.get(cell.vAlign);
                if (align) {
                    cellAttrs += ` align="${align}"`;
                }
                // A cell with all four edges takes the table's border; one with none has no border
                const spec = specs[i][j];
                if (border === 'borderCell' && spec === 0) {
                    cellAttrs += ' border="borderNone"';
                } else if (border === 'borderCell' && spec !== BORDER_ALL) {
                    cellAttrs += ` borderSpec="${spec}"`;
                }
                // The heaviest line the cell draws, unless it is Word's default line
                if (border === 'borderCell' && spec !== 0) {
                    const { line } = cellBorderChoice(cell, merge);
                    if (line && !isDefaultBorderLine(line)) {
                        const { borderStyle, borderWidth, borderColor } = cellBorderLine(line);
                        cellAttrs += ` borderStyle="${borderStyle}" borderWidth="${borderWidth}" borderColor="${borderColor}"`;
                    }
                }

                cellElements.push(`<cell${cellAttrs}>${cellContent}</cell>`);
            }

            rowElements.push(`<row rowName="row${i + 1}" rowType="dataRow">${cellElements.join('')}</row>`);
        }

        return `<table tableName="Table" columnCount="${columnCount}" columnSpans="${columnSpans}" border="${border}">${rowElements.join('')}</table>`;
    }

    /**
     * The <cell borderSpec> bits of the edges a parsed cell draws. A cell without parsed
     * edges draws all four, unless its table has no border.
     * @param {Object} cell - Parsed cell
     * @param {Object} table - Its parsed table
     * @returns {number} Edge bits (0 to 15)
     */
    cellEdges(cell, table) {
        const edges = cell.borders;
        if (!edges) return table.border === 'borderNone' ? 0 : BORDER_ALL;
        return (edges.top ? BORDER_TOP : 0) | (edges.right ? BORDER_RIGHT : 0) |
            (edges.bottom ? BORDER_BOTTOM : 0) | (edges.left ? BORDER_LEFT : 0);
    }

    /**
     * Content of a cell that holds no text: one paragraph covering no characters
     * @returns {string} Cell content elements
     */
    emptyCellContent() {
        return `<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="0" family="Times New Roman" size="12" /></paragraph>`;
    }

    /**
     * Process a table cell: its paragraphs, and the tables nested in it, in document order
     * @param {Object} cell - Parsed cell: its blocks (paragraphs and tables), else its paragraphs
     * @returns {string} Cell content elements
     */
    processTableCell(cell) {
        const cellElements = [];
        const blocks = cell.blocks || cell.paragraphs;

        for (let i = 0; i < blocks.length; i++) {
            // A nested table's cells add their own text; no line break follows the table
            if (blocks[i].type === 'table') {
                cellElements.push(this.tableElement(blocks[i]));
                continue;
            }

            const paragraph = blocks[i];
            const paraElements = [];

            for (const run of paragraph.runs) {
                if (run.type === 'text') {
                    const text = run.text;
                    const attrs = this.buildContentAttrs(run);
                    paraElements.push(`<content startOffset="${this.currentOffset}" length="${text.length}" ${attrs} />`);
                    this.content += text;
                    this.currentOffset += text.length;
                } else if (run.type === 'tab') {
                    paraElements.push(`<tab startOffset="${this.currentOffset}" length="1" />`);
                    this.content += '\t';
                    this.currentOffset += 1;
                } else if (run.type === 'break' || run.type === 'pageBreak') {
                    // A page break can't split a cell, so treat it as a line break
                    paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
                    this.content += '\n';
                    this.currentOffset += 1;
                } else if (run.type === 'image' && run.data) {
                    paraElements.push(`<image startOffset="${this.currentOffset}" length="1" imageData="${run.data}" width="${run.width}" height="${run.height}" />`);
                    this.content += '\uFFFC';
                    this.currentOffset += 1;
                } else if (run.type === 'footnoteRef' || run.type === 'endnoteRef') {
                    const mark = this.collectNote(run);
                    paraElements.push(`<content startOffset="${this.currentOffset}" length="${mark.length}" family="Times New Roman" size="10" superscript="true" />`);
                    this.content += mark;
                    this.currentOffset += mark.length;
                }
            }

            // Empty paragraph: still written, as every paragraph of a cell is (see paragraphIsWritten)
            if (paraElements.length === 0) {
                this.content += ' ';
                paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
                this.currentOffset += 1;
            }

            cellElements.push(this.buildParagraphElement(paragraph, paraElements));

            // Add newline between paragraphs, and before a nested table (except last)
            if (i < blocks.length - 1) {
                this.content += '\n';
                this.currentOffset += 1;
            }
        }

        // Empty cell
        if (cellElements.length === 0) {
            this.content += ' ';
            cellElements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" /></paragraph>`);
            this.currentOffset += 1;
        }

        return cellElements.join('');
    }

    /**
     * The <pageFormat> element from the parsed page setup. mediaSizeName stays 1 (A4),
     * the only paper size the UDF format documents.
     * @returns {string} The element
     */
    generatePageFormat() {
        const setup = this.pageSetup;
        const points = value => Number.isFinite(value) && value >= 0
            ? (Number.isInteger(value) ? value.toFixed(1) : String(value))
            : null;
        const attributes = {
            leftMargin: points(setup.leftMargin),
            rightMargin: points(setup.rightMargin),
            topMargin: points(setup.topMargin),
            bottomMargin: points(setup.bottomMargin),
            // 0 is landscape, 1 portrait
            paperOrientation: typeof setup.landscape === 'boolean' ? (setup.landscape ? '0' : '1') : null,
            headerFOffset: points(setup.headerOffset),
            footerFOffset: points(setup.footerOffset)
        };

        const values = Object.entries(DEFAULT_PAGE_FORMAT)
            .map(([name, fallback]) => `${name}="${attributes[name] ?? fallback}"`)
            .join(' ');
        return `<pageFormat mediaSizeName="1" ${values} />`;
    }

    /**
     * Generate final UDF XML
     * @returns {string} Complete XML content
     */
    generateXml() {
        const template = `<?xml version="1.0" encoding="UTF-8" ?>
<template format_id="1.8">
<content>${cdata(this.content)}</content>
<properties>${this.generatePageFormat()}</properties>
<elements resolver="hvl-default">
${this.elements.join('\n')}
</elements>
<styles><style name="default" description="Geçerli" family="Dialog" size="12" bold="false" italic="false" foreground="-13421773" FONT_ATTRIBUTE_KEY="javax.swing.plaf.FontUIResource[family=Dialog,name=Dialog,style=plain,size=12]" /><style name="hvl-default" family="Times New Roman" size="12" description="Gövde" /></styles>
</template>`;

        return template;
    }

    /**
     * Convert hex color to RGB integer (Java-style signed int)
     * @param {string} hex - Hex color (e.g., "#FF0000")
     * @returns {number} RGB integer
     */
    hexToRgbInt(hex) {
        return hexToRgbInt(hex); // The shared helper in utils.js
    }
}
