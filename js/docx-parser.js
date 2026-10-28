/**
 * DOCX Parser Module
 * Parses Word .docx files and extracts content with formatting
 */

// Namespaces an mc:Choice may require and still be read: text lives in w:p / w:r,
// text boxes in w:txbxContent and pictures in DrawingML
const SUPPORTED_CHOICE_NAMESPACES = [
    WORD_NS,
    DRAWING_NS,
    PICTURE_NS,
    WORDML_DRAWING_NS,
    'http://schemas.microsoft.com/office/word/2010/wordml',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
];

// Run formatting when neither styles.xml nor the run itself sets a property
const DEFAULT_RUN_FORMATTING = {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontFamily: 'Times New Roman',
    fontSize: 12
};

// On/off properties that toggle, rather than override, between the paragraph and
// character style levels (ECMA-376 Part 1, 17.7.3)
const TOGGLE_PROPERTIES = ['bold', 'italic', 'strike'];

// Points per unit of an OOXML universal measure (ST_UniversalMeasure, e.g. "2.5cm")
const POINTS_PER_UNIT = { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 };

class DocxParser {
    constructor() {
        this.zip = null;
        this.document = null;
        this.relationships = {};
        this.images = {};
        this.footnotes = {};
        this.numbering = {};
        this.styles = this.emptyStyles();
    }

    /**
     * Parse a DOCX file
     * @param {File} file - The DOCX file to parse
     * @returns {Promise<Object>} Parsed document structure
     */
    async parse(file) {
        try {
            // Reset state: the app reuses one parser for every file
            this.relationships = {};
            this.images = {};
            this.footnotes = {};
            this.numbering = {};
            this.styles = this.emptyStyles();

            // Load the DOCX as a ZIP
            this.zip = await JSZip.loadAsync(file);

            // Parse relationships to find media files
            await this.parseRelationships();

            // Extract images
            await this.extractImages();

            // Parse footnotes
            await this.parseFootnotes();

            // Parse list definitions
            await this.parseNumbering();

            // Parse style definitions
            await this.parseStyles();

            // Parse the main document content
            const documentXml = await this.zip.file('word/document.xml').async('string');
            this.document = parseXml(documentXml);

            // Extract document elements
            const elements = await this.extractElements();

            return {
                elements: elements,
                images: this.images,
                footnotes: this.footnotes,
                pageSetup: await this.parsePageSetup()
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
     * Parse list definitions from word/numbering.xml into
     * this.numbering[numId][ilvl] = { format, lvlText }
     */
    async parseNumbering() {
        const numberingFile = this.zip.file('word/numbering.xml');
        if (!numberingFile) return;

        try {
            const numberingDoc = parseXml(await numberingFile.async('string'));
            const root = numberingDoc.documentElement;

            const abstractLevels = {};
            for (const abstractNum of childElements(root, 'abstractNum')) {
                const levels = {};
                for (const lvl of childElements(abstractNum, 'lvl')) {
                    levels[lvl.getAttribute('w:ilvl') || '0'] = this.parseLevel(lvl);
                }
                abstractLevels[abstractNum.getAttribute('w:abstractNumId')] = levels;
            }

            for (const num of childElements(root, 'num')) {
                const abstractNumId = childElement(num, 'abstractNumId');
                const levels = { ...(abstractNumId && abstractLevels[abstractNumId.getAttribute('w:val')]) };

                // A level override replaces the whole abstract level, defaults included
                for (const override of childElements(num, 'lvlOverride')) {
                    const lvl = childElement(override, 'lvl');
                    if (lvl) levels[override.getAttribute('w:ilvl') || '0'] = this.parseLevel(lvl);
                }

                this.numbering[num.getAttribute('w:numId')] = levels;
            }
        } catch (error) {
            console.warn('Could not parse numbering:', error);
        }
    }

    /**
     * Page setup from the body's final w:sectPr, which describes the document's last
     * section: margins and header/footer offsets in points, and the orientation. Only
     * values the document sets, as valid numbers, are returned; the generator falls
     * back to its defaults for the rest.
     * @returns {Promise<{leftMargin?: number, rightMargin?: number, topMargin?: number,
     *   bottomMargin?: number, headerOffset?: number, footerOffset?: number, landscape?: boolean}>}
     */
    async parsePageSetup() {
        const pageSetup = {};
        const body = this.document.getElementsByTagNameNS(WORD_NS, 'body')[0];
        const sectPr = body && childElements(body, 'sectPr').pop();
        if (!sectPr) return pageSetup;

        const pgMar = childElement(sectPr, 'pgMar');
        if (pgMar) {
            const measure = attribute => this.measureToPoints(pgMar.getAttribute(attribute));
            const set = (key, value) => {
                if (value !== null) pageSetup[key] = value;
            };
            // A negative top or bottom margin is fixed (text may overlap the header or
            // footer); its size is the absolute value. The other values can't be negative.
            const absolute = value => value === null ? null : Math.abs(value);
            const nonNegative = value => value !== null && value >= 0 ? value : null;
            set('topMargin', absolute(measure('w:top')));
            set('bottomMargin', absolute(measure('w:bottom')));
            set('leftMargin', nonNegative(measure('w:left')));
            set('rightMargin', nonNegative(measure('w:right')));
            set('headerOffset', nonNegative(measure('w:header')));
            set('footerOffset', nonNegative(measure('w:footer')));

            // The gutter widens the binding-side margin: the top one with w:gutterAtTop
            // in settings.xml, the right one with w:rtlGutter, the left one otherwise
            const gutter = nonNegative(measure('w:gutter'));
            if (gutter) {
                const rtlGutter = childElement(sectPr, 'rtlGutter');
                const side = await this.isGutterAtTop() ? 'topMargin'
                    : rtlGutter && this.isOn(rtlGutter) ? 'rightMargin' : 'leftMargin';
                if (side in pageSetup) pageSetup[side] += gutter;
            }
        }

        const pgSz = childElement(sectPr, 'pgSz');
        if (pgSz) {
            const width = this.measureToPoints(pgSz.getAttribute('w:w'));
            const height = this.measureToPoints(pgSz.getAttribute('w:h'));
            pageSetup.landscape = pgSz.getAttribute('w:orient') === 'landscape' ||
                (width !== null && height !== null && width > height);
        }

        return pageSetup;
    }

    /**
     * Whether word/settings.xml puts the gutter at the top of the page (w:gutterAtTop)
     * @returns {Promise<boolean>}
     */
    async isGutterAtTop() {
        const settingsFile = this.zip.file('word/settings.xml');
        if (!settingsFile) return false;

        const root = parseXml(await settingsFile.async('string')).documentElement;
        const gutterAtTop = root.namespaceURI === WORD_NS && root.localName === 'settings'
            ? childElement(root, 'gutterAtTop') : null;
        return gutterAtTop ? this.isOn(gutterAtTop) : false;
    }

    /**
     * Convert an OOXML length to points: a number of twips (twentieths of a point),
     * or a universal measure with a unit such as "2.5cm"
     * @param {string|null} value - The attribute value
     * @returns {number|null} Points, or null for a missing or malformed value
     */
    measureToPoints(value) {
        const match = /^\s*(-?(?:\d+\.?\d*|\.\d+))(mm|cm|in|pt|pc|pi)?\s*$/.exec(value || '');
        if (!match) return null;
        const number = Number(match[1]);
        return match[2] ? number * POINTS_PER_UNIT[match[2]] : number / 20;
    }

    /**
     * Read the number format and level text of a numbering level
     * @param {Element} lvl - The w:lvl element
     * @returns {{format?: string, lvlText?: string}} Only the values the level sets
     */
    parseLevel(lvl) {
        const level = {};
        const numFmt = childElement(lvl, 'numFmt');
        if (numFmt && numFmt.hasAttribute('w:val')) level.format = numFmt.getAttribute('w:val');
        const lvlText = childElement(lvl, 'lvlText');
        if (lvlText && lvlText.hasAttribute('w:val')) level.lvlText = lvlText.getAttribute('w:val');
        return level;
    }

    /**
     * Style tables for a document without styles.xml: nothing set anywhere
     */
    emptyStyles() {
        return {
            defaults: { run: {}, paragraph: {} },
            definitions: { paragraph: new Map(), character: new Map() },
            defaultParagraph: null,
            resolved: { paragraph: new Map(), character: new Map() }
        };
    }

    /**
     * Parse word/styles.xml: the document defaults (w:docDefaults) and the paragraph
     * and character styles, each with the properties it sets itself. w:basedOn chains
     * are resolved on first use (see resolveStyle).
     */
    async parseStyles() {
        const stylesFile = this.zip.file('word/styles.xml');
        if (!stylesFile) return;

        try {
            const root = parseXml(await stylesFile.async('string')).documentElement;
            // A malformed part parses to a parsererror document instead of throwing
            if (root.namespaceURI !== WORD_NS || root.localName !== 'styles') {
                throw new Error('word/styles.xml is not a well-formed w:styles part');
            }

            const docDefaults = childElement(root, 'docDefaults');
            if (docDefaults) {
                const rPrDefault = childElement(docDefaults, 'rPrDefault');
                const pPrDefault = childElement(docDefaults, 'pPrDefault');
                this.styles.defaults = {
                    run: this.readRunProperties(rPrDefault && childElement(rPrDefault, 'rPr')),
                    paragraph: this.readParagraphProperties(pPrDefault && childElement(pPrDefault, 'pPr'))
                };
            }

            for (const style of childElements(root, 'style')) {
                const type = style.getAttribute('w:type');
                const definitions = this.styles.definitions[type];
                const id = style.getAttribute('w:styleId');
                if (!(type === 'paragraph' || type === 'character') || !id || definitions.has(id)) continue;

                const basedOn = childElement(style, 'basedOn');
                definitions.set(id, {
                    basedOn: basedOn ? basedOn.getAttribute('w:val') : null,
                    run: this.readRunProperties(childElement(style, 'rPr')),
                    paragraph: type === 'paragraph' ? this.readParagraphProperties(childElement(style, 'pPr')) : {}
                });

                if (type === 'paragraph' && this.styles.defaultParagraph === null && this.isOn(style, 'w:default', false)) {
                    this.styles.defaultParagraph = id;
                }
            }
        } catch (error) {
            this.styles = this.emptyStyles();
            console.warn('Could not parse styles:', error);
        }
    }

    /**
     * The properties a style gives, with its w:basedOn chain applied underneath it:
     * the nearest style in the chain that sets a property wins. A missing id, or a
     * chain that loops or names a missing style, stops the chain there.
     * @param {string} type - 'paragraph' or 'character'
     * @param {string|null} styleId - The style id (null: no style)
     * @returns {{run: Object, paragraph: Object}} Only the properties set somewhere
     */
    resolveStyle(type, styleId) {
        const resolved = this.styles.resolved[type];
        if (!resolved.has(styleId)) {
            const definitions = this.styles.definitions[type];
            const chain = [];
            for (let id = styleId; definitions.has(id) && !chain.includes(definitions.get(id)); id = definitions.get(id).basedOn) {
                chain.unshift(definitions.get(id));
            }

            resolved.set(styleId, {
                run: Object.assign({}, ...chain.map(style => style.run)),
                paragraph: Object.assign({}, ...chain.map(style => style.paragraph))
            });
        }
        return resolved.get(styleId);
    }

    /**
     * The resolved paragraph style of a paragraph: its w:pStyle, else the default paragraph style
     * @param {Element|null} pPr - The paragraph's w:pPr
     * @returns {{run: Object, paragraph: Object}} See resolveStyle
     */
    paragraphStyle(pPr) {
        const pStyle = pPr && childElement(pPr, 'pStyle');
        const id = pStyle && pStyle.getAttribute('w:val');
        return this.resolveStyle('paragraph',
            this.styles.definitions.paragraph.has(id) ? id : this.styles.defaultParagraph);
    }

    /**
     * Read the character properties a w:rPr sets, from a run, a style or the document defaults
     * @param {Element|null} rPr - The w:rPr element
     * @returns {Object} Only the properties it sets; an explicit "off" is kept as false,
     *   and color "auto" as null, so it overrides an earlier layer
     */
    readRunProperties(rPr) {
        const props = {};
        if (!rPr) return props;

        for (const [name, key] of [['b', 'bold'], ['i', 'italic'], ['strike', 'strike']]) {
            const el = childElement(rPr, name);
            if (el) props[key] = this.isOn(el);
        }

        const u = childElement(rPr, 'u');
        if (u && u.hasAttribute('w:val')) props.underline = u.getAttribute('w:val') !== 'none';

        const rFonts = childElement(rPr, 'rFonts');
        if (rFonts && rFonts.getAttribute('w:ascii')) props.fontFamily = rFonts.getAttribute('w:ascii');

        // Font size is in half-points
        const sz = childElement(rPr, 'sz');
        const halfPoints = sz ? parseInt(sz.getAttribute('w:val')) : NaN;
        if (halfPoints > 0) props.fontSize = halfPoints / 2;

        const color = childElement(rPr, 'color');
        const colorVal = color && color.getAttribute('w:val');
        if (colorVal) props.color = colorVal === 'auto' ? null : '#' + colorVal;

        return props;
    }

    /**
     * Read the alignment and indents a w:pPr sets, from a paragraph, a style or the document defaults
     * @param {Element|null} pPr - The w:pPr element
     * @returns {Object} Only the properties it sets
     */
    readParagraphProperties(pPr) {
        const props = {};
        if (!pPr) return props;

        const jc = childElement(pPr, 'jc');
        if (jc) props.alignment = this.mapAlignment(jc.getAttribute('w:val'));

        const ind = childElement(pPr, 'ind');
        if (ind) {
            if (ind.hasAttribute('w:left')) props.leftIndent = this.twipsToPoints(ind.getAttribute('w:left'));
            if (ind.hasAttribute('w:right')) props.rightIndent = this.twipsToPoints(ind.getAttribute('w:right'));
            if (ind.hasAttribute('w:firstLine')) {
                props.firstLineIndent = this.twipsToPoints(ind.getAttribute('w:firstLine'));
            } else if (ind.hasAttribute('w:hanging')) {
                // A hanging indent can't be represented, but it still replaces a first-line indent
                props.firstLineIndent = 0;
            }
        }

        return props;
    }

    /**
     * Whether an on/off property is on: absent w:val means on, "false", "0" and "off" mean off
     * @param {Element} el - The property element
     * @param {string} [attribute] - The attribute holding the value
     * @param {boolean} [ifAbsent] - The value when the attribute is absent
     * @returns {boolean}
     */
    isOn(el, attribute = 'w:val', ifAbsent = true) {
        if (!el.hasAttribute(attribute)) return ifAbsent;
        return !['false', '0', 'off'].includes(el.getAttribute(attribute));
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
     * content controls (w:sdt), custom XML wrappers and one branch of mc:AlternateContent
     * @param {Element} container - The containing element
     * @returns {Array<Element>} Word ML children in document order
     */
    contentChildren(container) {
        const children = [];

        for (const child of container.children) {
            if (child.namespaceURI === MARKUP_COMPATIBILITY_NS && child.localName === 'AlternateContent') {
                const branch = this.alternateContentBranch(child);
                if (branch) children.push(...this.contentChildren(branch));
            } else if (child.namespaceURI !== WORD_NS) {
                continue;
            } else if (child.localName === 'customXml') {
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
     * Pick the one branch of mc:AlternateContent to read: the first mc:Choice whose
     * required namespaces this parser understands, else mc:Fallback. Reading both
     * would duplicate content.
     * @param {Element} alternateContent - The mc:AlternateContent element
     * @returns {Element|null} The branch to read
     */
    alternateContentBranch(alternateContent) {
        const choices = childElements(alternateContent, 'Choice', MARKUP_COMPATIBILITY_NS);
        const supported = choices.find(choice =>
            (choice.getAttribute('Requires') || '').split(/\s+/).filter(Boolean)
                .every(prefix => SUPPORTED_CHOICE_NAMESPACES.includes(choice.lookupNamespaceURI(prefix))));

        // With no usable choice and no fallback, a choice is still better than nothing
        return supported || childElement(alternateContent, 'Fallback', MARKUP_COMPATIBILITY_NS) || choices[0] || null;
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

        // Alignment and indents: document defaults, the paragraph style chain, then the paragraph's own w:pPr
        const pPr = childElement(para, 'pPr');
        const style = this.paragraphStyle(pPr);
        Object.assign(paragraph, this.styles.defaults.paragraph, style.paragraph, this.readParagraphProperties(pPr));

        if (pPr) {
            // Numbering
            const numPr = childElement(pPr, 'numPr');
            if (numPr) {
                const ilvl = childElement(numPr, 'ilvl');
                const numId = childElement(numPr, 'numId');
                // numId 0 switches numbering off for the paragraph
                if (ilvl && numId && numId.getAttribute('w:val') !== '0') {
                    paragraph.numbering = {
                        level: parseInt(ilvl.getAttribute('w:val') || '0'),
                        numId: numId.getAttribute('w:val')
                    };
                    // The list definition, if numbering.xml has one for this level
                    const definition = (this.numbering[paragraph.numbering.numId] || {})[paragraph.numbering.level];
                    if (definition) {
                        // A level without w:numFmt is decimal
                        paragraph.numbering.format = definition.format || 'decimal';
                        paragraph.numbering.lvlText = definition.lvlText || '';
                    }
                }
            }
        }

        // Get runs (text with formatting)
        for (const run of runs) {
            paragraph.runs.push(...this.parseRun(run, style.run));
        }

        return paragraph;
    }

    /**
     * Parse a run element (text with formatting)
     * @param {Element} run - The run element
     * @param {Object} [paragraphRun] - Run properties from the paragraph's style (see resolveStyle)
     * @returns {Array<Object>} Parsed items in document order (text, tab, break, pageBreak, image, footnoteRef)
     */
    parseRun(run, paragraphRun = {}) {
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
                    if (!formatting) formatting = this.parseRunFormatting(run, paragraphRun);
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
                if (drawing) {
                    items.push(this.parseDrawing(drawing));
                }
            }
        }

        return items;
    }

    /**
     * Resolve a run's character formatting in Word's order: built-in defaults, document
     * defaults, the paragraph style chain, the character style chain (w:rStyle), then
     * the run's own w:rPr. Later layers override earlier ones, except that bold, italic
     * and strike set by the two style levels toggle: on in both styles is off, and off
     * in a style leaves the value unchanged (unless the document defaults turn it on).
     * Direct formatting always sets them outright.
     * @param {Element} run - The run element
     * @param {Object} [paragraphRun] - Run properties from the paragraph's style (see resolveStyle)
     * @returns {Object} Text item template (without text)
     */
    parseRunFormatting(run, paragraphRun = {}) {
        const rPr = childElement(run, 'rPr');
        const rStyle = rPr && childElement(rPr, 'rStyle');
        const defaults = this.styles.defaults.run;
        const characterRun = rStyle ? this.resolveStyle('character', rStyle.getAttribute('w:val')).run : {};
        const direct = this.readRunProperties(rPr);
        const formatting = {
            type: 'text',
            ...DEFAULT_RUN_FORMATTING,
            ...defaults,
            ...paragraphRun,
            ...characterRun,
            ...direct
        };

        for (const key of TOGGLE_PROPERTIES) {
            if (key in direct) continue;
            if (defaults[key] === true) {
                // On in the document defaults: a style level that leaves it unset counts as on, and
                // Word shows it on when an even number of levels turn it off (MS-OI29500, 17.7.3)
                formatting[key] = [paragraphRun[key], characterRun[key]].filter(value => value === false).length % 2 === 0;
            } else {
                formatting[key] = (paragraphRun[key] === true) !== (characterRun[key] === true);
            }
        }
        if (!formatting.color) delete formatting.color;

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

        // Get embedded image reference. Pictures inside a text box belong to the
        // text box's own paragraphs, which are parsed separately.
        const blips = Array.from(drawing.getElementsByTagNameNS(DRAWING_NS, 'blip'))
            .filter(blip => !this.isInTextBox(blip, drawing));
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
     * Whether an element sits inside a text box (w:txbxContent) below an ancestor
     * @param {Element} element - The element to check
     * @param {Element} ancestor - Where to stop looking
     * @returns {boolean}
     */
    isInTextBox(element, ancestor) {
        for (let el = element.parentElement; el && el !== ancestor; el = el.parentElement) {
            if (el.namespaceURI === WORD_NS && el.localName === 'txbxContent') return true;
        }
        return false;
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
