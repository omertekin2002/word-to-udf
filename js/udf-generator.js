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

class UdfGenerator {
    constructor() {
        this.content = '';
        this.pageSetup = {};
        this.elements = [];
        this.currentOffset = 0;
        this.collectedFootnotes = []; // Store footnotes to append at end
        this.footnoteCounter = 0; // Track display numbers for footnotes
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
        this.pageSetup = document.pageSetup || {};

        // Process all document elements
        for (const element of document.elements) {
            if (element.type === 'paragraph') {
                this.processParagraph(element);
            } else if (element.type === 'table') {
                this.processTable(element);
            }
        }

        // Append footnotes section at the end if there are any
        if (this.collectedFootnotes.length > 0) {
            this.appendFootnotesSection();
        }

        // Ensure we have at least one paragraph
        if (this.content.length === 0) {
            this.content = '\u200B'; // Zero-width space
            this.elements.push('<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="0" length="1" family="Times New Roman" size="12" /></paragraph>');
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
                // Handle page break - first flush any pending footnotes
                this.content += paraContent.join('');
                const emittedFragment = paraElements.length > 0;
                if (emittedFragment) {
                    this.elements.push(this.buildParagraphElement(paraProps, paraElements));
                }

                // Insert footnotes for this page before the page break
                if (this.collectedFootnotes.length > 0) {
                    this.appendFootnotesSection();
                    this.collectedFootnotes = []; // Clear for next page
                }

                this.elements.push('<page-break />');

                // Runs after the break continue in a new paragraph on the next page.
                // It is the same Word paragraph, so it must not start another list
                // item or repeat the first-line indent.
                paraContent.length = 0;
                paraElements.length = 0;
                afterPageBreak = true;
                if (emittedFragment) {
                    paraProps = { ...paragraph, numbering: null, firstLineIndent: 0 };
                }
            } else if (run.type === 'footnoteRef') {
                // Handle footnote reference - insert superscript number only
                // Use counter for display number (w:id can be non-contiguous)
                this.footnoteCounter++;
                const displayNum = String(this.footnoteCounter);
                const footnoteText = run.content;

                // Collect footnote for later
                this.collectedFootnotes.push({
                    number: displayNum,
                    text: footnoteText
                });

                // Add the superscript footnote number in text
                paraElements.push(`<content startOffset="${this.currentOffset}" length="${displayNum.length}" family="Times New Roman" size="10" superscript="true" />`);
                paraContent.push(displayNum);
                this.currentOffset += displayNum.length;
            }
        }

        // If paragraph is empty, add placeholder (unless it only held a page break)
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
     * Append footnotes section at the end of the document
     */
    appendFootnotesSection() {
        // Add a separator line
        const separator = '────────────────────';

        // Empty line before separator (element references this newline)
        this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" /></paragraph>`);
        this.content += '\n';
        this.currentOffset += 1;

        // Separator line
        this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0"><content startOffset="${this.currentOffset}" length="${separator.length}" family="Times New Roman" size="10" /></paragraph>`);
        this.content += separator;
        this.currentOffset += separator.length;

        // Add each footnote
        for (const footnote of this.collectedFootnotes) {
            // Create paragraph with superscript number and regular text
            const numPart = footnote.number + '. ';
            const textPart = footnote.text;

            const paraElements = [];

            // Superscript number
            paraElements.push(`<content startOffset="${this.currentOffset}" length="${numPart.length}" family="Times New Roman" size="10" superscript="true" />`);
            this.content += numPart;
            this.currentOffset += numPart.length;

            // Footnote text
            paraElements.push(`<content startOffset="${this.currentOffset}" length="${textPart.length}" family="Times New Roman" size="10" />`);
            this.content += textPart;
            this.currentOffset += textPart.length;

            this.elements.push(`<paragraph Alignment="0" LeftIndent="0.0" RightIndent="0.0">${paraElements.join('')}</paragraph>`);
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

        if (run.color) {
            // Convert hex color to RGB integer
            const rgb = this.hexToRgbInt(run.color);
            attrs.push(`foreground="${rgb}"`);
        }

        return attrs.join(' ');
    }

    /**
     * Process a table element
     * @param {Object} table - Parsed table
     */
    processTable(table) {
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

        const rowElements = [];

        for (let i = 0; i < table.rows.length; i++) {
            const row = table.rows[i];
            const cellElements = [];

            for (const cell of row.cells) {
                if (cell.vMergeContinue) continue; // Skip merged cells

                const cellContent = this.processTableCell(cell);

                let cellAttrs = '';
                if (cell.colspan > 1) {
                    cellAttrs += ` colspan="${cell.colspan}"`;
                }
                if (cell.bgColor) {
                    const rgb = this.hexToRgbInt(cell.bgColor);
                    cellAttrs += ` bgColor="${rgb}"`;
                }
                if (cell.vAlign) {
                    cellAttrs += ` vAlign="${cell.vAlign}"`;
                }

                cellElements.push(`<cell${cellAttrs}>${cellContent}</cell>`);
            }

            rowElements.push(`<row rowName="row${i + 1}" rowType="dataRow">${cellElements.join('')}</row>`);
        }

        const tableElement = `<table tableName="Table" columnCount="${columnCount}" columnSpans="${columnSpans}" border="${table.border}">${rowElements.join('')}</table>`;
        this.elements.push(tableElement);
    }

    /**
     * Process a table cell
     * @param {Object} cell - Parsed cell
     * @returns {string} Cell content elements
     */
    processTableCell(cell) {
        const cellElements = [];

        for (let i = 0; i < cell.paragraphs.length; i++) {
            const paragraph = cell.paragraphs[i];
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
                } else if (run.type === 'footnoteRef') {
                    // Handle footnote reference in table cells
                    // Use counter for display number (w:id can be non-contiguous)
                    this.footnoteCounter++;
                    const displayNum = String(this.footnoteCounter);
                    const footnoteText = run.content;

                    // Collect footnote for later
                    this.collectedFootnotes.push({
                        number: displayNum,
                        text: footnoteText
                    });

                    // Add the superscript footnote number
                    paraElements.push(`<content startOffset="${this.currentOffset}" length="${displayNum.length}" family="Times New Roman" size="10" superscript="true" />`);
                    this.content += displayNum;
                    this.currentOffset += displayNum.length;
                }
            }

            // Empty paragraph
            if (paraElements.length === 0) {
                this.content += ' ';
                paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
                this.currentOffset += 1;
            }

            cellElements.push(this.buildParagraphElement(paragraph, paraElements));

            // Add newline between paragraphs (except last)
            if (i < cell.paragraphs.length - 1) {
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
        if (!hex) return -16777216; // Black

        hex = hex.replace('#', '');
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);

        // Java-style signed RGB (0xFFRRGGBB as signed 32-bit)
        const rgb = (255 << 24) | (r << 16) | (g << 8) | b;
        return rgb | 0; // Convert to signed 32-bit
    }
}
