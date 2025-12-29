/**
 * UDF Generator Module
 * Generates UYAP UDF format from parsed document structure
 */

class UdfGenerator {
    constructor() {
        this.content = '';
        this.elements = [];
        this.currentOffset = 0;
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

        // Process all document elements
        for (const element of document.elements) {
            if (element.type === 'paragraph') {
                this.processParagraph(element);
            } else if (element.type === 'table') {
                this.processTable(element);
            }
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
                paraContent.push('\n');
                this.currentOffset += 1;
            } else if (run.type === 'pageBreak') {
                // Handle page break
                this.content += paraContent.join('');
                if (paraElements.length > 0) {
                    this.elements.push(this.buildParagraphElement(paragraph, paraElements));
                }
                this.elements.push('<page-break />');
                return;
            }
        }

        // If paragraph is empty, add placeholder
        if (paraContent.length === 0) {
            paraContent.push('\u200B'); // Zero-width space
            paraElements.push(`<content startOffset="${this.currentOffset}" length="1" family="Times New Roman" size="12" />`);
            this.currentOffset += 1;
        }

        this.content += paraContent.join('');
        this.elements.push(this.buildParagraphElement(paragraph, paraElements));
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
            const level = paragraph.numbering.level;
            const numId = paragraph.numbering.numId;

            // Determine if it's numbered or bulleted based on numId
            // (simplified - in real implementation, would check numbering definitions)
            const isBulleted = parseInt(numId) % 2 === 0;

            if (isBulleted) {
                attrs += ` Bulleted="true" BulletType="BULLET_TYPE_ELLIPSE" ListLevel="${level}" ListId="${numId}"`;
            } else {
                attrs += ` Numbered="true" NumberType="NUMBER_TYPE_NUMBER_TRE" ListLevel="${level}" ListId="${numId}"`;
            }
        }

        return `<paragraph ${attrs}>${childElements.join('')}</paragraph>`;
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
                } else if (run.type === 'image' && run.data) {
                    paraElements.push(`<image startOffset="${this.currentOffset}" length="1" imageData="${run.data}" width="${run.width}" height="${run.height}" />`);
                    this.content += '\uFFFC';
                    this.currentOffset += 1;
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
     * Generate final UDF XML
     * @returns {string} Complete XML content
     */
    generateXml() {
        const template = `<?xml version="1.0" encoding="UTF-8" ?>
<template format_id="1.8">
<content><![CDATA[${this.content}]]></content>
<properties><pageFormat mediaSizeName="1" leftMargin="42.51968479156494" rightMargin="28.34645652770996" topMargin="14.17322826385498" bottomMargin="14.17322826385498" paperOrientation="1" headerFOffset="20.0" footerFOffset="20.0" /></properties>
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
