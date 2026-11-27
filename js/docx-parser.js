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
    hidden: false,
    caps: false,
    smallCaps: false,
    fontFamily: 'Times New Roman',
    fontSize: 12
};

// The theme font slots a w:rFonts w:asciiTheme may name for Latin text, and the theme font
// (a:majorFont or a:minorFont, ECMA-376 Part 1, 20.1.4.1.24 / 20.1.4.1.25) each resolves to
const THEME_FONT_SLOTS = { majorAscii: 'major', majorHAnsi: 'major', minorAscii: 'minor', minorHAnsi: 'minor' };

// Word's fixed highlighting palette (w:highlight w:val, ST_HighlightColor, ECMA-376 Part 1, 17.18.40)
const HIGHLIGHT_COLORS = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
    darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF'
};

// On/off properties that toggle, rather than override, between the paragraph and
// character style levels (ECMA-376 Part 1, 17.7.3). Double strikethrough (w:dstrike) isn't
// one: unlike w:strike, the standard doesn't call it a toggle property (17.3.2.9), so a later
// layer simply overrides it.
const TOGGLE_PROPERTIES = ['bold', 'italic', 'strike', 'hidden', 'caps', 'smallCaps'];

// Run children that hold a picture, text box, chart or shape: DrawingML w:drawing, and the
// VML w:pict and w:object (an OLE object, shown by its preview picture)
const PICTURE_ELEMENTS = ['drawing', 'pict', 'object'];

// Run children that show something, so a hidden run (w:vanish) holding one leaves content out.
// w:t counts only when it has text; a picture may also be wrapped in mc:AlternateContent.
const HIDDEN_CONTENT_ELEMENTS = ['tab', 'ptab', 'br', 'cr', 'noBreakHyphen', 'sym', 'footnoteReference', 'endnoteReference',
    ...PICTURE_ELEMENTS];

// Width and height of a picture whose size the document doesn't give, in points
const DEFAULT_IMAGE_SIZE = 100;

// Points per unit of a CSS length in a VML shape's style (e.g. "width:120pt"); a number
// without a unit is in pixels, at 96 per inch
const VML_POINTS_PER_UNIT = { pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, px: 72 / 96 };

// Points per unit of an OOXML universal measure (ST_UniversalMeasure, e.g. "2.5cm")
const POINTS_PER_UNIT = { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 };

// Note types (w:type) of a footnote or endnote that are Word's separator lines, not note text
const SEPARATOR_NOTE_TYPES = ['separator', 'continuationSeparator', 'continuationNotice'];

// Run children of a footnote or endnote that stand for whitespace: a tab or a line break
const NOTE_WHITESPACE_ELEMENTS = ['tab', 'ptab', 'br', 'cr'];

// Parsed run items a note paragraph keeps: page breaks, pictures and note references have no place there
const NOTE_RUN_TYPES = ['text', 'tab', 'break'];

// Font size of note text, in points, when neither the document defaults nor a style or the run gives one
const NOTE_FONT_SIZE = 10;

// Unicode text of the symbols Word writes as w:sym (Insert > Symbol, and check boxes), by
// lower-case font name and character code. Word writes w:char as hex, usually in the private
// use range F020-F0FF and sometimes as the plain 0020-00FF; both are looked up by the low byte.
// A code a table doesn't list has no equivalent and is left out (see symbolText).
const SYMBOL_FONTS = {
    // Symbol: every code in Unicode's MAPPINGS/VENDORS/ADOBE/symbol.txt (Adobe Symbol Encoding
    // to Unicode, table version 1.0, 2011-07-12), except the 29 it maps into the Corporate Use
    // Subarea (bracket, brace and integral pieces, arrow extenders and the serif and sans-serif
    // registered, copyright and trade mark signs). Where the file gives a code two values, the
    // first one listed is used, except 0x6D, written as GREEK SMALL LETTER MU rather than MICRO SIGN.
    // The angle brackets (0xE1, 0xF1) are escaped: normalizing the file would turn them into U+3008 and U+3009.
    'symbol': {
        0x20: ' ', 0x21: '!', 0x22: '∀', 0x23: '#', 0x24: '∃', 0x25: '%', 0x26: '&', 0x27: '∋',
        0x28: '(', 0x29: ')', 0x2A: '∗', 0x2B: '+', 0x2C: ',', 0x2D: '−', 0x2E: '.', 0x2F: '/',
        0x30: '0', 0x31: '1', 0x32: '2', 0x33: '3', 0x34: '4', 0x35: '5', 0x36: '6', 0x37: '7',
        0x38: '8', 0x39: '9', 0x3A: ':', 0x3B: ';', 0x3C: '<', 0x3D: '=', 0x3E: '>', 0x3F: '?',
        0x40: '≅', 0x41: 'Α', 0x42: 'Β', 0x43: 'Χ', 0x44: 'Δ', 0x45: 'Ε', 0x46: 'Φ', 0x47: 'Γ',
        0x48: 'Η', 0x49: 'Ι', 0x4A: 'ϑ', 0x4B: 'Κ', 0x4C: 'Λ', 0x4D: 'Μ', 0x4E: 'Ν', 0x4F: 'Ο',
        0x50: 'Π', 0x51: 'Θ', 0x52: 'Ρ', 0x53: 'Σ', 0x54: 'Τ', 0x55: 'Υ', 0x56: 'ς', 0x57: 'Ω',
        0x58: 'Ξ', 0x59: 'Ψ', 0x5A: 'Ζ', 0x5B: '[', 0x5C: '∴', 0x5D: ']', 0x5E: '⊥', 0x5F: '_',
        0x61: 'α', 0x62: 'β', 0x63: 'χ', 0x64: 'δ', 0x65: 'ε', 0x66: 'φ', 0x67: 'γ', 0x68: 'η',
        0x69: 'ι', 0x6A: 'ϕ', 0x6B: 'κ', 0x6C: 'λ', 0x6D: 'μ', 0x6E: 'ν', 0x6F: 'ο', 0x70: 'π',
        0x71: 'θ', 0x72: 'ρ', 0x73: 'σ', 0x74: 'τ', 0x75: 'υ', 0x76: 'ϖ', 0x77: 'ω', 0x78: 'ξ',
        0x79: 'ψ', 0x7A: 'ζ', 0x7B: '{', 0x7C: '|', 0x7D: '}', 0x7E: '∼', 0xA0: '€', 0xA1: 'ϒ',
        0xA2: '′', 0xA3: '≤', 0xA4: '⁄', 0xA5: '∞', 0xA6: 'ƒ', 0xA7: '♣', 0xA8: '♦', 0xA9: '♥',
        0xAA: '♠', 0xAB: '↔', 0xAC: '←', 0xAD: '↑', 0xAE: '→', 0xAF: '↓', 0xB0: '°', 0xB1: '±',
        0xB2: '″', 0xB3: '≥', 0xB4: '×', 0xB5: '∝', 0xB6: '∂', 0xB7: '•', 0xB8: '÷', 0xB9: '≠',
        0xBA: '≡', 0xBB: '≈', 0xBC: '…', 0xBF: '↵', 0xC0: 'ℵ', 0xC1: 'ℑ', 0xC2: 'ℜ', 0xC3: '℘',
        0xC4: '⊗', 0xC5: '⊕', 0xC6: '∅', 0xC7: '∩', 0xC8: '∪', 0xC9: '⊃', 0xCA: '⊇', 0xCB: '⊄',
        0xCC: '⊂', 0xCD: '⊆', 0xCE: '∈', 0xCF: '∉', 0xD0: '∠', 0xD1: '∇', 0xD5: '∏', 0xD6: '√',
        0xD7: '⋅', 0xD8: '¬', 0xD9: '∧', 0xDA: '∨', 0xDB: '⇔', 0xDC: '⇐', 0xDD: '⇑', 0xDE: '⇒',
        0xDF: '⇓', 0xE0: '◊', 0xE1: '\u2329', 0xE5: '∑', 0xF1: '\u232A', 0xF2: '∫', 0xF3: '⌠', 0xF5: '⌡'
    },
    // Wingdings and Wingdings 2: Unicode document L2/11-052R, "Wingdings and Webdings Symbols"
    // (Michel Suignard, 2011-02-15; linked from the README), whose glyph IDs w-1nnn (Wingdings) and w-2nnn (Wingdings 2) end in the decimal character code.
    // A glyph it unifies with an encoded character (its UCS column) maps to that character; a
    // variant maps to the character its "→" cross-reference names (marked "ref"). Only the
    // boxes, check and cross marks, bullets and arrows petitions use are listed. The glyphs the
    // document gives neither for, such as Wingdings 2's empty box (A3, w-2163 BOX 1), are unmapped.
    'wingdings': {
        0x6C: '●', // w-1108 BLACK CIRCLE 6, ref U+25CF
        0x6E: '■', // w-1110 BLACK SQUARE 6, ref U+25A0
        0x6F: '□', // w-1111 BOX 3, ref U+25A1
        0x71: '❑', // w-1113 U+2751 LOWER RIGHT SHADOWED WHITE SQUARE
        0x72: '❒', // w-1114 U+2752 UPPER RIGHT SHADOWED WHITE SQUARE
        0x75: '◆', // w-1117 BLACK DIAMOND 6, ref U+25C6
        0x76: '❖', // w-1118 U+2756 BLACK DIAMOND MINUS WHITE X
        0x77: '⬩', // w-1119 BLACK DIAMOND 4, ref U+2B29
        0x78: '☒', // w-1120 U+2612 BALLOT BOX WITH X
        0x9E: '⋅', // w-1158 BLACK CIRCLE 2, ref U+22C5
        0xA7: '▪', // w-1167 BLACK SQUARE 4, ref U+25AA
        0xA8: '□', // w-1168 BOX 2, ref U+25A1
        0xD8: '➢', // w-1216 THREE-D TOP-LIGHTED RIGHTWARDS ARROWHEAD, ref U+27A2
        0xE8: '➔', // w-1232 WIDE-HEADED RIGHTWARDS ARROW 4, ref U+2794
        0xEF: '⇦', // w-1239 U+21E6 LEFTWARDS WHITE ARROW
        0xF0: '⇨', // w-1240 U+21E8 RIGHTWARDS WHITE ARROW
        0xF1: '⇧', // w-1241 U+21E7 UPWARDS WHITE ARROW
        0xF2: '⇩', // w-1242 U+21E9 DOWNWARDS WHITE ARROW
        0xF3: '⬄', // w-1243 U+2B04 LEFT RIGHT WHITE ARROW
        0xF4: '⇳', // w-1244 U+21F3 UP DOWN WHITE ARROW
        0xF5: '⬁', // w-1245 U+2B01 NORTH WEST WHITE ARROW
        0xF6: '⬀', // w-1246 U+2B00 NORTH EAST WHITE ARROW
        0xF7: '⬃', // w-1247 U+2B03 SOUTH WEST WHITE ARROW
        0xF8: '⬂', // w-1248 U+2B02 SOUTH EAST WHITE ARROW
        0xFB: '✗', // w-1251 BALLOT X, ref U+2717
        0xFC: '✓', // w-1252 U+2713 CHECK MARK
        0xFD: '☒', // w-1253 BALLOT BOX WITH X, ref U+2612
        0xFE: '☑' // w-1254 U+2611 BALLOT BOX WITH CHECK
    },
    'wingdings 2': {
        0x4F: '✗', // w-2079 BALLOT X, see w-1251: ref U+2717
        0x50: '✓', // w-2080 U+2713 CHECK MARK
        0x51: '☒', // w-2081 BALLOT BOX WITH X, see w-1253: ref U+2612
        0x52: '☑', // w-2082 U+2611 BALLOT BOX WITH CHECK
        0x53: '☒', // w-2083 BALLOT BOX WITH LIGHT X, ref U+2612
        0x54: '☒', // w-2084 U+2612 BALLOT BOX WITH X
        0x55: '⊗' // w-2085 CIRCLED X, ref U+2297
    }
};

// Endnote number formats (w:endnotePr/w:numFmt) the generator can write. Word numbers
// endnotes in lower-case Roman numerals unless the document says otherwise.
const ENDNOTE_FORMATS = ['decimal', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter'];
const DEFAULT_ENDNOTE_FORMAT = 'lowerRoman';

// Signature of an OLE compound file: a Word 97-2003 .doc, or a password-protected .docx
// UDF TabSet alignment codes for Word's w:tab w:val (left, center, right, decimal).
// A num stop is a left stop; a bar stop draws a line and positions no text, so it isn't listed.
const TAB_ALIGNMENTS = { left: 0, start: 0, num: 0, center: 1, right: 2, end: 2, decimal: 3 };

// UDF TabSet leader codes for Word's w:tab w:leader; anything else (none, underscore, heavy) is no leader
const TAB_LEADERS = { dot: 1, middleDot: 1, hyphen: 2 };

// Word leaders that are drawn, but have no UDF code, so they become no leader
const LOST_TAB_LEADERS = ['underscore', 'heavy'];

// Decimal places of a point that tell two tab stop positions apart (see tabStopKey)
const TAB_POSITION_PRECISION = 6;

// Line rules (w:spacing w:lineRule) lineSpacing() can't express, so their paragraphs get single spacing
const LOST_LINE_RULES = ['exact', 'atLeast'];

// Run items a header or footer paragraph keeps: page breaks and note references have no place there
const HEADER_FOOTER_RUN_TYPES = ['text', 'tab', 'break', 'image'];

// Start of the VML shape id Word gives a watermark it inserts in a header (Design › Watermark):
// a picture watermark, and a text one (a WordArt shape). The UDF's page background (<bgImage>)
// isn't written, and a watermark isn't a picture in the header's text.
const WATERMARK_SHAPE_IDS = /^(?:WordPictureWatermark|PowerPlusWaterMarkObject)/;

// Table edges a w:tblBorders sets, each under the names Word writes for it (start and end are
// the newer names of left and right)
const TABLE_EDGES = {
    top: ['top'],
    left: ['left', 'start'],
    bottom: ['bottom'],
    right: ['right', 'end'],
    insideH: ['insideH'],
    insideV: ['insideV']
};

// Cell edges a w:tcBorders sets
const CELL_EDGES = {
    top: TABLE_EDGES.top,
    left: TABLE_EDGES.left,
    bottom: TABLE_EDGES.bottom,
    right: TABLE_EDGES.right
};

// Border values (ST_Border) that draw no line
const NO_BORDER_VALUES = ['none', 'nil'];

// Text of an equation part that needs no parentheses around it in linear form (see mathOperand):
// one run of letters or digits, or a decimal number
const SIMPLE_MATH_OPERAND = /^(?:[\p{L}\p{N}]+|\p{N}+(?:[.,]\p{N}+)+)$/u;

// Word ML elements inside an equation that show no text: run properties, and deleted or moved-away runs
const MATH_SKIPPED_WORD_ELEMENTS = ['rPr', 'del', 'moveFrom'];

const COMPOUND_FILE_SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

// Name of the stream an encrypted OOXML package keeps its encryption parameters in (MS-OFFCRYPTO 2.3.4)
const ENCRYPTION_INFO_STREAM = 'EncryptionInfo';

// Compound file sector number that ends a chain, and the "no entry" directory ID (MS-CFB 2.1)
const COMPOUND_END_OF_CHAIN = 0xFFFFFFFE;
const COMPOUND_NO_STREAM = 0xFFFFFFFF;

// Relationship from the package to its main part: transitional, then Strict OOXML
const OFFICE_DOCUMENT_RELATIONSHIPS = [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
];

// Relationship types from the main document part to the supporting parts read here
const PART_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

// WordprocessingML namespace of a document saved as Strict Open XML (ECMA-376 Part 1, Strict)
const STRICT_WORD_NS = 'http://purl.oclc.org/ooxml/wordprocessingml/main';

// Messages for the parse failures the app explains; error.code tells them apart
const PARSE_FAILURES = {
    corrupt: 'The Word document is damaged: its main document part is missing or not well-formed.',
    encrypted: 'The Word document is password-protected.',
    embeddedOnly: 'The Word document has no text of its own: its content is embedded by another program (w:altChunk).',
    legacyDoc: 'The file is a Word 97-2003 document (.doc), not a .docx.',
    strict: 'The Word document is saved as Strict Open XML, which is not supported.',
    tooLarge: 'The Word document is too large to convert: its parts decompress past the size limit.'
};

// Most bytes one part of the .docx may decompress to, and all the parts one parse reads
// together. A petition is far below both: they only stop a zip bomb, or a damaged size
// field, from running the tab out of memory.
const MAX_PART_BYTES = 100 * 1024 * 1024;
const MAX_TOTAL_BYTES = 300 * 1024 * 1024;

// Codes of the conversion findings parse() reports, in the order it lists them: content the
// conversion drops or simplifies, so the user can fix it in Word before filing
const FINDING_CODES = ['headersFooters', 'comments', 'hiddenText', 'nestedTables', 'missingImage', 'unsupportedImage',
    'unconvertedDrawing', 'notesMoved', 'endnotesMoved', 'lineSpacing', 'tabStops', 'symbols', 'tableBorders',
    'equations', 'embeddedContent', 'trackedChanges', 'sectionLayout', 'subscript', 'doubleStrikethrough',
    'listMarkers', 'paragraphBorders', 'hyperlinks'];

// The scheme a link's address starts with (http://, https://, ftp://, ..., or mailto:), left off its
// hyperlinks example and when checking whether the link's text shows the address
const LINK_SCHEME = /^(?:[a-z][a-z0-9+.-]*:\/\/|mailto:)/i;

// Switches of a HYPERLINK field instruction that take an argument (ECMA-376 Part 1, 17.16.5.25): the
// bookmark (\l), the tooltip (\o) and the target frame (\t). The others (\m, \n) take none.
const HYPERLINK_SWITCHES_WITH_ARGUMENT = ['\\l', '\\o', '\\t'];

// Vertical alignments a run can have (w:vertAlign w:val, ST_VerticalAlignRun, ECMA-376 Part 1, 17.18.96).
// Superscript is written; subscript has no UDF attribute, so it stays on the baseline and is reported.
const VERTICAL_ALIGNMENTS = ['baseline', 'superscript', 'subscript'];

// The edges of a paragraph's border (w:pBdr, ECMA-376 Part 1, 17.3.1.24): the box's four sides,
// the line between paragraphs with the same borders, and a bar beside it
const PARAGRAPH_BORDER_EDGES = ['top', 'left', 'bottom', 'right', 'between', 'bar'];

// Border line styles (w:val) that draw nothing
const INVISIBLE_BORDER_STYLES = ['nil', 'none'];

// Section start types (a w:sectPr's w:type, which says how that section starts after the previous
// one) that start the section on a new page. A missing w:type means nextPage. oddPage and evenPage
// are written as a plain page break and reported (sectionLayout): Word may add a blank page there.
const PAGE_SECTION_STARTS = ['nextPage', 'oddPage', 'evenPage'];
const PARITY_SECTION_STARTS = ['oddPage', 'evenPage'];

// The w:pgSz and w:pgMar attributes a section's page setup is compared on (see sectionLayoutDiffers)
const PAGE_SIZE_ATTRIBUTES = ['w:w', 'w:h'];
const PAGE_MARGIN_ATTRIBUTES = ['w:top', 'w:right', 'w:bottom', 'w:left'];

// Track Changes revision marks inside a paragraph, its runs and its paragraph mark: inserted, deleted
// and moved content (w:ins, w:del, w:moveTo, w:moveFrom, which also mark a paragraph mark in
// w:pPr/w:rPr and inserted numbering in w:numPr), and formatting, numbering and section changes
const PARAGRAPH_REVISION_ELEMENTS = ['ins', 'del', 'moveTo', 'moveFrom', 'rPrChange', 'pPrChange', 'numberingChange',
    'sectPrChange'];
// In a row's w:trPr: an inserted or deleted row, or a row formatting change
const ROW_REVISION_ELEMENTS = ['ins', 'del', 'trPrChange'];
// In a cell's w:tcPr: an inserted, deleted or merged cell, or a cell formatting change
const CELL_REVISION_ELEMENTS = ['cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];
// In w:pPr/w:rPr: a deleted or moved-away paragraph mark, which Word's Accept All removes
const DELETED_MARK_ELEMENTS = ['del', 'moveFrom'];
// Inline wrappers whose runs a document with every change accepted doesn't have
const DELETED_RUN_WRAPPERS = ['del', 'moveFrom'];

// How many location hints a finding keeps, and how long each may be
const MAX_FINDING_EXAMPLES = 3;
const MAX_EXAMPLE_LENGTH = 40;

// How many located occurrences (an example and where it is) a finding keeps
const MAX_FINDING_OCCURRENCES = 5;

// How many levels of tables are written as tables, the body's own table included: by default a
// table in a cell of a table in a cell of a body table is the deepest. A table nested deeper is
// flattened into its cell's paragraphs and reported (nestedTables). 1 flattens every nested table.
const MAX_TABLE_DEPTH = 3;

// Outline levels (w:outlineLvl w:val) that make a paragraph a heading: 0 to 8 are Word's
// heading levels 1 to 9, and 9 is body text
const MAX_HEADING_OUTLINE_LEVEL = 8;

// Longest text a paragraph without an outline level may have and still count as a heading
// when it is all bold or all upper case (see isPetitionHeading)
const MAX_PETITION_HEADING_LENGTH = 80;

// Picture formats, told apart by their leading bytes rather than the part's file extension.
// UDF documents only PNG and JPEG. EMF starts with an EMR_HEADER record (type 1) whose
// signature " EMF" sits at byte 40; WMF either starts with a placeable header or with a
// METAHEADER of type 1 or 2 and header size 9.
const IMAGE_SIGNATURES = [
    ['PNG', [0x89, 0x50, 0x4E, 0x47]],
    ['JPEG', [0xFF, 0xD8, 0xFF]],
    ['GIF', [0x47, 0x49, 0x46, 0x38]],
    ['TIFF', [0x49, 0x49, 0x2A, 0x00]],
    ['TIFF', [0x4D, 0x4D, 0x00, 0x2A]],
    ['WMF', [0xD7, 0xCD, 0xC6, 0x9A]],
    ['WMF', [0x01, 0x00, 0x09, 0x00]],
    ['WMF', [0x02, 0x00, 0x09, 0x00]]
];
const SUPPORTED_IMAGE_FORMATS = ['PNG', 'JPEG'];

/**
 * Name of a picture's format from its leading bytes
 * @param {Uint8Array} bytes - The picture part
 * @returns {string} 'PNG', 'JPEG', 'EMF', 'WMF', 'GIF', 'TIFF', 'SVG' or 'unknown'
 */
function imageFormat(bytes) {
    const startsWith = (signature, offset = 0) => signature.every((byte, i) => bytes[offset + i] === byte);
    if (startsWith([0x01, 0x00, 0x00, 0x00]) && startsWith([0x20, 0x45, 0x4D, 0x46], 40)) return 'EMF';
    const match = IMAGE_SIGNATURES.find(([, signature]) => startsWith(signature));
    if (match) return match[0];
    // SVG is XML text: look for its root element near the start, past any prolog or comment
    const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 1024));
    if (/^﻿?\s*</.test(head) && /<svg[\s>]/.test(head)) return 'SVG';
    return 'unknown';
}

/**
 * A short location hint from document text: whitespace collapsed, cut to MAX_EXAMPLE_LENGTH
 * @param {string} text
 * @returns {string} '' for blank text
 */
function shortExample(text) {
    const collapsed = text.replace(/\s+/g, ' ').trim();
    return collapsed.length > MAX_EXAMPLE_LENGTH ? `${collapsed.slice(0, MAX_EXAMPLE_LENGTH).trimEnd()}…` : collapsed;
}

/**
 * The address a HYPERLINK field instruction links to: its first argument that is neither a switch
 * nor a switch's argument, e.g. the address in HYPERLINK "<address>" \o "tip"
 * @param {string} instruction - The field instruction (w:instrText of a complex field, w:instr of a simple one)
 * @returns {string} '' for another field type, or a link to a bookmark only (HYPERLINK \l "name")
 */
function hyperlinkFieldAddress(instruction) {
    const match = /^\s*HYPERLINK(?=[\s"]|$)([\s\S]*)$/i.exec(instruction);
    if (!match) return '';
    const tokens = Array.from(match[1].matchAll(/"([^"]*)"|(\S+)/g), ([, quoted, bare]) => ({ quoted: quoted !== undefined, text: quoted ?? bare }));
    for (let i = 0; i < tokens.length; i++) {
        const { quoted, text } = tokens[i];
        if (!quoted && text.startsWith('\\')) {
            if (HYPERLINK_SWITCHES_WITH_ARGUMENT.includes(text.toLowerCase())) i++;
            continue;
        }
        if (text.trim() !== '') return text.trim();
    }
    return '';
}

/**
 * Whether a link's text already shows its address: it contains the address, compared ignoring case,
 * a leading scheme (see LINK_SCHEME), a leading www. and trailing slashes, so a text showing
 * www.<domain> shows the address https://www.<domain>/
 * @param {string} text - The link's shown text
 * @param {string} address - Where it links to
 * @returns {boolean}
 */
function linkShowsAddress(text, address) {
    const core = address.trim().replace(LINK_SCHEME, '').replace(/^www\./i, '').replace(/\/+$/, '').toLowerCase();
    return text.toLowerCase().includes(core);
}

/**
 * Unicode text of a symbol (w:sym), from SYMBOL_FONTS, and its symbols finding example: the
 * font and code, then an arrow and the text it became. A symbol left out has no arrow.
 * @param {Element} sym - The w:sym element
 * @param {string} [runFont] - The run's font, used when w:sym names none
 * @returns {{text: string, example: string}} text is '' for a symbol with no equivalent: an
 *   unknown font, or a code its font's table doesn't list
 */
function symbolText(sym, runFont = '') {
    const font = (sym.getAttribute('w:font') || runFont || '').trim();
    const char = (sym.getAttribute('w:char') || '').trim().toUpperCase();
    const code = /^[0-9A-F]{1,4}$/.test(char) ? parseInt(char, 16) : -1;
    const text = symbolFontCharacter(font, code);
    const label = `${font || '?'} ${char || '?'}`;
    return { text, example: shortExample(text ? `${label} → ${text}` : label) };
}

/**
 * Unicode text of a character code in a symbol font, from SYMBOL_FONTS
 * @param {string} font - The font name
 * @param {number} code - The character code: F0xx, the private use form, or 00xx, the plain one
 * @returns {string} '' when the font has no table, or its table doesn't list the code
 */
function symbolFontCharacter(font, code) {
    const table = SYMBOL_FONTS[font.trim().toLowerCase()];
    const highByte = code >> 8;
    return table && code >= 0 && (highByte === 0xF0 || highByte === 0) &&
        Object.prototype.hasOwnProperty.call(table, code & 0xFF) ? table[code & 0xFF] : '';
}

/**
 * The character a bullet list level draws: its level text (w:lvlText), read through its font
 * (w:lvl/w:rPr/w:rFonts) when that is a symbol font in SYMBOL_FONTS, as for w:sym. A level text
 * of one character a symbol font's table lists becomes that table's character; one in the private
 * use form (F0xx) the table doesn't list has none; any other level text is drawn as written.
 * @param {string} lvlText - The level text
 * @param {string} font - The level's font, or ''
 * @returns {string} The bullet as Unicode text, '' when it has no equivalent
 */
function bulletCharacter(lvlText, font) {
    const characters = Array.from(lvlText);
    if (!SYMBOL_FONTS[font.trim().toLowerCase()] || characters.length !== 1) return lvlText;
    const code = characters[0].codePointAt(0);
    const text = symbolFontCharacter(font, code);
    if (text) return text;
    return code >> 8 === 0xF0 ? '' : lvlText;
}

/**
 * The whole number in an element's w:val, such as w:start's
 * @param {Element|null} element
 * @returns {number|null} null without the element, or for a w:val that isn't a whole number
 */
function integerValue(element) {
    const value = element ? (element.getAttribute('w:val') || '').trim() : '';
    return /^-?\d+$/.test(value) ? Number(value) : null;
}

/**
 * What a hidden run (w:vanish) would show: its text, and whether it holds anything Word would
 * show, text or one of HIDDEN_CONTENT_ELEMENTS, a drawing included
 * @param {Element} run - The w:r element
 * @returns {{shows: boolean, text: string}}
 */
function hiddenRunContent(run) {
    const text = childElements(run, 't').map(t => t.textContent || '').join('');
    const shows = text !== '' || Array.from(run.children).some(child =>
        HIDDEN_CONTENT_ELEMENTS.includes(child.localName) ||
        (child.localName !== 'rPr' && PICTURE_ELEMENTS.some(name => child.getElementsByTagNameNS(WORD_NS, name).length > 0)));
    return { shows, text };
}

/**
 * A length in a VML element's style attribute, e.g. "width:120pt"
 * @param {Element} element - A VML element
 * @param {string} name - 'width' or 'height'
 * @returns {{value: number, unit: string}|null} The number and its lower-case unit ('' for a
 *   bare number), or null when the style doesn't give the length or gives it unreadably
 */
function vmlStyleLength(element, name) {
    for (const declaration of ((element && element.getAttribute('style')) || '').split(';')) {
        const colon = declaration.indexOf(':');
        if (colon < 0 || declaration.slice(0, colon).trim().toLowerCase() !== name) continue;
        const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)([a-z]*)$/i.exec(declaration.slice(colon + 1).trim());
        return match ? { value: parseFloat(match[1]), unit: match[2].toLowerCase() } : null;
    }
    return null;
}

/**
 * Points of a VML length. A bare number in a shape inside a v:group is in the group's
 * coordinate system (coordsize, 1000 by default), which the group's own width or height
 * spans; elsewhere it is in pixels.
 * @param {Element} element - The VML element the length belongs to
 * @param {string} name - 'width' or 'height'
 * @param {{value: number, unit: string}} length
 * @returns {number|null} Points, or null for an unknown unit or a group whose size is unreadable
 */
function vmlLengthToPoints(element, name, { value, unit }) {
    const group = element.parentElement;
    if (!unit && group && group.namespaceURI === VML_NS && group.localName === 'group') {
        const groupLength = vmlStyleLength(group, name);
        const coordinates = (group.getAttribute('coordsize') || '').split(',').map(part => parseFloat(part));
        const span = coordinates[name === 'width' ? 0 : 1];
        const coordinateSize = Number.isFinite(span) ? span : 1000;
        if (!groupLength || !(coordinateSize > 0)) return null;
        return vmlLengthToPoints(group, name, { value: value * groupLength.value / coordinateSize, unit: groupLength.unit });
    }
    const key = unit || 'px';
    return Object.prototype.hasOwnProperty.call(VML_POINTS_PER_UNIT, key) ? value * VML_POINTS_PER_UNIT[key] : null;
}

/**
 * Size of a VML shape from the width and height in its style attribute, scaled through any
 * v:group holding it
 * @param {Element} shape - The v:shape (or other VML element) that owns a v:imagedata
 * @returns {{width: number, height: number}} Rounded points; DEFAULT_IMAGE_SIZE for a side
 *   that is missing, unreadable, non-finite or not positive
 */
function vmlShapeSize(shape) {
    const size = { width: DEFAULT_IMAGE_SIZE, height: DEFAULT_IMAGE_SIZE };
    for (const name of ['width', 'height']) {
        const length = vmlStyleLength(shape, name);
        const points = length && vmlLengthToPoints(shape, name, length);
        const rounded = points === null ? NaN : Math.round(points);
        if (Number.isFinite(rounded) && rounded > 0) size[name] = rounded;
    }
    return size;
}

/**
 * Whether a parsed body paragraph reads as a petition's heading though no outline level makes it
 * one: non-empty text of at most MAX_PETITION_HEADING_LENGTH characters in which every text run
 * with visible text is bold, or every letter is upper case ("AÇIKLAMALAR :", "SONUÇ VE İSTEM")
 * @param {Object} paragraph - A parsed paragraph
 * @returns {boolean}
 */
function isPetitionHeading(paragraph) {
    const text = paragraphText(paragraph).trim();
    if (text === '' || text.length > MAX_PETITION_HEADING_LENGTH) return false;
    const visibleRuns = paragraph.runs.filter(run => run.type === 'text' && run.text.trim() !== '');
    if (visibleRuns.every(run => run.bold)) return true;
    return /\p{L}/u.test(text) && !/\p{Ll}/u.test(text);
}

/**
 * Whether an Office Math on/off property is on: present, with no m:val or one that isn't off
 * @param {Element|null} property - e.g. m:degHide
 * @returns {boolean}
 */
function mathIsOn(property) {
    if (!property) return false;
    if (!property.hasAttributeNS(MATH_NS, 'val')) return true;
    return !['0', 'false', 'off'].includes(property.getAttributeNS(MATH_NS, 'val'));
}

/**
 * A character an Office Math structure's properties set, e.g. m:dPr/m:begChr
 * @param {Element} element - The structure, e.g. m:d
 * @param {string} properties - Local name of its properties element, e.g. 'dPr'
 * @param {string} name - Local name of the character property, e.g. 'begChr'
 * @param {string} fallback - The character when the property or its m:val is missing
 * @returns {string} '' when m:val is empty: no character
 */
function mathCharacter(element, properties, name, fallback) {
    const pr = childElement(element, properties, MATH_NS);
    const property = pr && childElement(pr, name, MATH_NS);
    return property && property.hasAttributeNS(MATH_NS, 'val') ? property.getAttributeNS(MATH_NS, 'val') : fallback;
}

/**
 * Whether an element is inside an equation (m:oMath) below an ancestor
 * @param {Element} element
 * @param {Element} ancestor - Where to stop looking
 * @returns {boolean}
 */
function isInEquation(element, ancestor) {
    for (let el = element.parentElement; el && el !== ancestor; el = el.parentElement) {
        if (el.namespaceURI === MATH_NS && el.localName === 'oMath') return true;
    }
    return false;
}

/**
 * A copy of a finding's location, so the caller can't change the one the parser keeps
 * @param {Object} location - See DocxParser.addFinding
 * @returns {Object}
 */
function copyLocation(location) {
    const copy = { ...location };
    if (location.note) copy.note = { ...location.note };
    return copy;
}

/**
 * The w:sectPr of a section break: the one in a paragraph's w:pPr, which ends a section there
 * @param {Element} para - The w:p element
 * @returns {Element|null}
 */
function sectionBreakOf(para) {
    const pPr = childElement(para, 'pPr');
    return (pPr && childElement(pPr, 'sectPr')) || null;
}

/**
 * How a section starts after the previous one: its w:type, nextPage when it has none
 * (the w:type default). A missing final w:sectPr is a section with nothing set.
 * @param {Element|null} sectPr - The section's w:sectPr
 * @returns {string} e.g. 'nextPage', 'continuous', 'oddPage'
 */
function sectionStart(sectPr) {
    const type = sectPr && childElement(sectPr, 'type');
    return (type && type.getAttribute('w:val')) || 'nextPage';
}

/**
 * Whether a section has more than one text column: w:cols with w:num above 1, or more than one w:col
 * @param {Element} sectPr - The section's w:sectPr
 * @returns {boolean}
 */
function hasColumns(sectPr) {
    const cols = childElement(sectPr, 'cols');
    if (!cols) return false;
    return parseInt(cols.getAttribute('w:num'), 10) > 1 || childElements(cols, 'col').length > 1;
}

/**
 * Plain text of a parsed paragraph's text runs
 * @param {Object} paragraph - A parsed paragraph
 * @returns {string}
 */
function paragraphText(paragraph) {
    return paragraph.runs.filter(run => run.type === 'text').map(run => run.text).join('');
}

/**
 * Whether a paragraph's mark is deleted or moved away (w:pPr/w:rPr/w:del or w:moveFrom), so
 * Word's Accept All joins it to the next paragraph
 * @param {Element} para - The w:p element
 * @returns {boolean}
 */
function paragraphMarkDeleted(para) {
    const pPr = childElement(para, 'pPr');
    const rPr = pPr && childElement(pPr, 'rPr');
    return Boolean(rPr) && DELETED_MARK_ELEMENTS.some(name => childElement(rPr, name));
}

/**
 * Whether Accept All removes this row and its contents (w:trPr/w:del)
 */
function tableRowDeleted(trPr) {
    return Boolean(trPr && childElement(trPr, 'del'));
}

/**
 * Whether a parsed paragraph has nothing left: no items, or only whitespace text. With its mark
 * deleted too (see paragraphMarkDeleted), Word's Accept All removes it.
 * @param {Array<Object>} items - Its parsed items (see parseRun)
 * @returns {boolean}
 */
function nothingLeft(items) {
    return items.every(item => item.type === 'text' && item.text.trim() === '');
}

/**
 * Whether an element is inside deleted or moved-away content (w:del, w:moveFrom) of a paragraph
 * @param {Element} element - A run
 * @param {Element} ancestor - The paragraph to stop at
 * @returns {boolean}
 */
function isDeletedRun(element, ancestor) {
    for (let el = element.parentElement; el && el !== ancestor; el = el.parentElement) {
        if (el.namespaceURI === WORD_NS && DELETED_RUN_WRAPPERS.includes(el.localName)) return true;
    }
    return false;
}

/**
 * Whether a parsed run item shows something in the UDF: text other than whitespace, a picture
 * with data, or a note reference
 * @param {Object} item - See parseRun
 * @returns {boolean}
 */
function showsContent(item) {
    if (item.type === 'text') return item.text.trim() !== '';
    if (item.type === 'image') return Boolean(item.data);
    return item.type === 'footnoteRef' || item.type === 'endnoteRef';
}

/**
 * The identity of a tab stop position: stops are at the same position when their keys are
 * equal. A length converts to slightly different numbers of points depending on its unit
 * (432 twips is 21.6, 0.3in is 21.599999999999998), so positions are compared to a millionth
 * of a point. That is far finer than a twip (0.05 points), the closest two stops Word keeps.
 * @param {number} position - Points, as measureToPoints gives them
 * @returns {string}
 */
function tabStopKey(position) {
    return position.toFixed(TAB_POSITION_PRECISION);
}

/**
 * An Error for a parse failure the app explains to the user
 * @param {string} code - A PARSE_FAILURES key
 * @returns {Error}
 */
function parseFailure(code) {
    const error = new Error(PARSE_FAILURES[code]);
    error.code = code;
    return error;
}

/**
 * Name of the part a relationship targets, without a leading slash: an absolute target
 * from the package root, a relative one from the folder of the part the relationship
 * belongs to (ECMA-376 Part 2, 9.3)
 * @param {string} sourcePart - The part whose relationship it is ('' for the package)
 * @param {string} target - The relationship's Target
 * @returns {string}
 */
function resolvePartName(sourcePart, target) {
    const segments = target.startsWith('/') ? [] : sourcePart.split('/').slice(0, -1);
    for (const segment of target.split('/')) {
        if (segment === '..') segments.pop();
        else if (segment && segment !== '.') segments.push(segment);
    }
    return segments.join('/');
}

/**
 * Read the bytes of a parse() input
 * @param {Blob|ArrayBuffer|Uint8Array} file - A browser File, or a Buffer in the tests
 * @param {number} [length] - Read only this many leading bytes
 * @returns {Promise<Uint8Array|null>} null for an input that isn't one of those
 */
async function readBytes(file, length) {
    if (file && typeof file.arrayBuffer === 'function' && typeof file.slice === 'function') {
        const blob = length === undefined ? file : file.slice(0, length);
        return new Uint8Array(await blob.arrayBuffer());
    }
    let bytes = null;
    if (ArrayBuffer.isView(file)) {
        bytes = new Uint8Array(file.buffer, file.byteOffset, file.byteLength);
    } else if (Object.prototype.toString.call(file) === '[object ArrayBuffer]') {
        bytes = new Uint8Array(file);
    }
    return bytes && length !== undefined ? bytes.subarray(0, length) : bytes;
}

/**
 * A parse failure that a part read raises: an optional part's reader must not swallow it,
 * since it says the whole package can't be read safely
 * @param {string} code - 'tooLarge', or 'corrupt' for a part longer than its declared size
 * @returns {Error}
 */
function partReadFailure(code) {
    const error = parseFailure(code);
    error.fatal = true;
    return error;
}

/**
 * Rethrow a failure no optional part's reader may swallow (see partReadFailure)
 * @param {*} error - What a part's reader caught
 */
function rethrowFatal(error) {
    if (error && error.fatal) throw error;
}

/**
 * Uncompressed size a ZIP entry declares in its central directory record, which JSZip
 * 3.10.1 keeps on the entry's compressed data. Read defensively: it's not public API.
 * JSZip reads the 32-bit field as a signed integer, so a size of 2 GiB or more comes back
 * negative and is turned back into the unsigned value.
 * @param {JSZipObject} file - The part
 * @returns {number|null} null when the entry doesn't say
 */
function declaredSize(file) {
    const size = file && file._data && file._data.uncompressedSize;
    return Number.isInteger(size) && size >= -0x80000000 && size <= 0xFFFFFFFF ? size >>> 0 : null;
}

/**
 * Decompress a part through JSZip's stream API, stopping as soon as it produces more bytes
 * than it declares ('corrupt') or than the limit allows ('tooLarge'), so a size field that
 * lies can't make it allocate past the limit
 * @param {JSZipObject} file - The part
 * @param {number} limit - Most bytes the part may decompress to
 * @param {number|null} declared - Its declared uncompressed size, if known
 * @returns {Promise<Uint8Array>}
 */
function inflatePart(file, limit, declared) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let length = 0;
        let settled = false;
        let stream;
        const fail = error => {
            if (settled) return;
            settled = true;
            if (stream) stream.pause();
            reject(error);
        };
        try {
            stream = file.internalStream('uint8array');
        } catch (error) {
            fail(error);
            return;
        }
        stream.on('data', chunk => {
            // A paused stream still delivers what the chunk being decompressed produces
            if (settled) return;
            length += chunk.length;
            if (declared !== null && length > declared) return fail(partReadFailure('corrupt'));
            if (length > limit) return fail(partReadFailure('tooLarge'));
            chunks.push(chunk);
        }).on('error', fail).on('end', () => {
            if (settled) return;
            settled = true;
            const bytes = new Uint8Array(length);
            let offset = 0;
            for (const chunk of chunks) {
                bytes.set(chunk, offset);
                offset += chunk.length;
            }
            resolve(bytes);
        });
        stream.resume();
    });
}

/**
 * Decode an XML part's bytes as text. A package part may be UTF-8 or UTF-16 (ECMA-376
 * Part 2), told apart by its byte order mark or, without one, by how the leading "<" is encoded
 * @param {Uint8Array} bytes - The part
 * @returns {string}
 */
function decodeXmlPart(bytes) {
    let encoding = 'utf-8';
    if ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0x3C && bytes[1] === 0x00)) {
        encoding = 'utf-16le';
    } else if ((bytes[0] === 0xFE && bytes[1] === 0xFF) || (bytes[0] === 0x00 && bytes[1] === 0x3C)) {
        encoding = 'utf-16be';
    }
    return new TextDecoder(encoding).decode(bytes);
}

/**
 * Names of the streams directly in a compound file's root storage (MS-CFB). The directory
 * is read by following its sector chain through the FAT, and only the root's children are
 * visited, so a stream nested in a storage (an embedded object) or in another stream's
 * bytes (an embedded file) doesn't count.
 * @param {Uint8Array} bytes - The whole compound file
 * @returns {Set<string>} Empty for a malformed file
 */
function compoundRootStreams(bytes) {
    const names = new Set();
    if (bytes.length < 512) return names;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const uint32 = offset => (offset + 4 <= bytes.length ? view.getUint32(offset, true) : COMPOUND_NO_STREAM);
    const sectorShift = view.getUint16(0x1E, true);
    if (sectorShift !== 9 && sectorShift !== 12) return names;
    const sectorSize = 1 << sectorShift;
    const sectorCount = Math.ceil(bytes.length / sectorSize);
    // The header fills sector -1, so sector n starts one sector further in
    const sectorOffset = sector => (sector + 1) * sectorSize;

    // FAT sector numbers: the first 109 are in the header, the rest in the DIFAT sector chain
    const fatSectorCount = Math.min(uint32(0x2C), sectorCount);
    const fatSectors = [];
    for (let i = 0; i < 109 && fatSectors.length < fatSectorCount; i++) fatSectors.push(uint32(0x4C + i * 4));
    for (let difat = uint32(0x44), hops = 0; fatSectors.length < fatSectorCount && difat < sectorCount &&
        hops < sectorCount; difat = uint32(sectorOffset(difat) + sectorSize - 4), hops++) {
        for (let i = 0; i < sectorSize / 4 - 1 && fatSectors.length < fatSectorCount; i++) {
            fatSectors.push(uint32(sectorOffset(difat) + i * 4));
        }
    }
    const perFatSector = sectorSize / 4;
    const nextSector = sector => {
        const fatSector = fatSectors[Math.floor(sector / perFatSector)];
        return fatSector === undefined ? COMPOUND_END_OF_CHAIN
            : uint32(sectorOffset(fatSector) + (sector % perFatSector) * 4);
    };

    // Directory entries are 128 bytes; entry n is the nth in the directory chain
    const entries = [];
    for (let sector = uint32(0x30), hops = 0; sector < sectorCount && hops < sectorCount;
        sector = nextSector(sector), hops++) {
        for (let offset = 0; offset < sectorSize; offset += 128) entries.push(sectorOffset(sector) + offset);
    }
    const isEntry = id => entries[id] !== undefined && entries[id] + 128 <= bytes.length;
    if (!isEntry(0) || bytes[entries[0] + 66] !== 5) return names;

    // The root's children form a tree linked by left (68) and right (72) sibling IDs
    const pending = [uint32(entries[0] + 76)];
    const seen = new Set();
    while (pending.length > 0) {
        const id = pending.pop();
        if (!isEntry(id) || seen.has(id)) continue;
        seen.add(id);
        const entry = entries[id];
        pending.push(uint32(entry + 68), uint32(entry + 72));
        // A stream (type 2) with a UTF-16LE name of nameLength bytes, terminator included
        const nameLength = view.getUint16(entry + 64, true);
        if (bytes[entry + 66] !== 2 || nameLength < 2 || nameLength > 64) continue;
        let name = '';
        for (let i = 0; i < nameLength / 2 - 1; i++) name += String.fromCharCode(view.getUint16(entry + i * 2, true));
        names.add(name);
    }
    return names;
}

class DocxParser {
    constructor() {
        this.zip = null;
        this.documentPath = 'word/document.xml';
        this.document = null;
        this.relationships = {};
        // The relationships of the header or footer part being parsed, which its pictures
        // resolve through (see embedPicture); null while the main document part is parsed
        this.partRelationships = null;
        this.images = {};
        this.imageFormats = {};
        this.footnotes = {};
        this.endnotes = {};
        // The [code, example] findings of each note's text, recorded when a reference to it is converted
        this.noteFindings = { footnote: {}, endnote: {} };
        // The formatted paragraphs of each note (see parseNoteParagraphs), keyed like footnotes and endnotes
        this.noteParagraphs = { footnote: {}, endnote: {} };
        this.numbering = {};
        // The theme's Latin typefaces, which w:asciiTheme names (see parseThemeFonts)
        this.themeFonts = { major: null, minor: null };
        this.styles = this.emptyStyles();
        this.findings = {};
        this.inDiscardedCell = false;
        // How many tables deep the content being parsed is (see parseTable): 0 outside any table
        this.tableDepth = 0;
        // Whether a note's paragraphs are being parsed (see parseNoteParagraphs): addFinding records
        // nothing, since noteText's findings are recorded once per reference instead, and parseRun
        // leaves out pictures and note references
        this.parsingNote = false;
        // Where findings are recorded now (see addFinding), and the body's paragraph and table counts
        // and last heading so far (see parseBodyBlocks)
        this.location = { kind: 'document' };
        this.body = { paragraphs: 0, tables: 0, headings: { outline: null, petition: null } };
        this.documentParagraphs = null;
        // Each parsed paragraph's resolved paragraph style id and whether it has contextual
        // spacing (see applyContextualSpacing)
        this.contextualSpacing = new WeakMap();
        this.endnotesMoved = false;
        // Decompression limits, in bytes; tests lower them instead of building huge files
        this.limits = { part: MAX_PART_BYTES, total: MAX_TOTAL_BYTES };
        this.partBytes = new Map();
        this.bytesRead = 0;
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
            this.partRelationships = null;
            this.images = {};
            this.imageFormats = {};
            this.footnotes = {};
            this.endnotes = {};
            this.noteFindings = { footnote: {}, endnote: {} };
            this.noteParagraphs = { footnote: {}, endnote: {} };
            this.numbering = {};
            this.themeFonts = { major: null, minor: null };
            this.styles = this.emptyStyles();
            this.findings = {};
            this.inDiscardedCell = false;
            this.tableDepth = 0;
            this.parsingNote = false;
            this.location = { kind: 'document' };
            this.body = { paragraphs: 0, tables: 0, headings: { outline: null, petition: null } };
            this.documentParagraphs = null;
            this.contextualSpacing = new WeakMap();
            this.endnotesMoved = false;
            this.embeddedChunks = 0;
            this.partBytes = new Map();
            this.bytesRead = 0;

            // A compound file is not a ZIP: name why it can't be read instead of a generic failure
            await this.rejectCompoundFile(file);

            // Load the DOCX as a ZIP
            this.zip = await JSZip.loadAsync(file);
            this.documentPath = await this.findMainDocument();

            // Parse relationships to find media files
            await this.parseRelationships();

            // Parse list definitions
            await this.parseNumbering();

            // Read the theme's fonts, which styles and runs name through w:asciiTheme
            await this.parseThemeFonts();

            // Parse style definitions
            await this.parseStyles();

            // Parse footnotes and endnotes, after the styles their runs' hidden formatting resolves through
            this.footnotes = await this.parseNotes('footnote');
            this.endnotes = await this.parseNotes('endnote');

            // Parse the main document content. A malformed part parses to a parsererror
            // document instead of throwing (or, in Chromium, keeps a partial w:document with a
            // parsererror inside), and would convert to a blank UDF
            const documentFile = this.zip.file(this.documentPath);
            if (!documentFile) throw parseFailure('corrupt');
            this.document = parseXml(await this.readXmlPart(documentFile));
            const root = this.document.documentElement;
            if (!root || this.document.getElementsByTagName('parsererror').length > 0) {
                throw parseFailure('corrupt');
            }
            // A well-formed Strict Open XML document is not damaged, just in a namespace not read here
            if (root.namespaceURI === STRICT_WORD_NS && root.localName === 'document') {
                throw parseFailure('strict');
            }
            if (root.namespaceURI !== WORD_NS || root.localName !== 'document') {
                throw parseFailure('corrupt');
            }

            // The final section's default header and footer are converted; other header and
            // footer text, and comments, are not, and are noted
            const headersFooters = await this.parseHeadersFooters();
            await this.detectComments();

            // Whether converting moves endnotes Word places per section or renumbers;
            // each converted endnote reference then gets a finding
            this.endnotesMoved = await this.endnotesAreMoved();

            // Decode the pictures the document shows, then extract document elements
            await this.extractImages();
            const elements = await this.extractElements();

            // A body whose only content is embedded chunks (w:altChunk), which aren't converted,
            // would give an empty UDF
            if (this.embeddedChunks > 0 && !this.flattenToParagraphs(elements).some(paragraph =>
                paragraph.runs.some(showsContent))) {
                throw parseFailure('embeddedOnly');
            }

            return {
                elements: elements,
                header: headersFooters.header,
                footer: headersFooters.footer,
                images: this.images,
                footnotes: this.footnotes,
                endnotes: this.endnotes,
                endnoteFormat: await this.parseEndnoteFormat(),
                pageSetup: await this.parsePageSetup(),
                findings: this.listFindings()
            };
        } catch (error) {
            console.error('Error parsing DOCX:', error);
            if (Object.prototype.hasOwnProperty.call(PARSE_FAILURES, error && error.code)) {
                throw parseFailure(error.code);
            }
            throw new Error('Failed to parse DOCX file. Make sure it is a valid Word document.');
        }
    }

    /**
     * Throw if the file is an OLE compound file: 'encrypted' when its root storage holds
     * an EncryptionInfo stream (a password-protected .docx), 'legacyDoc' otherwise
     * @param {Blob|ArrayBuffer|Uint8Array} file - The input given to parse()
     */
    async rejectCompoundFile(file) {
        const header = await readBytes(file, COMPOUND_FILE_SIGNATURE.length);
        if (!header || header.length < COMPOUND_FILE_SIGNATURE.length ||
            !COMPOUND_FILE_SIGNATURE.every((byte, i) => header[i] === byte)) {
            return;
        }
        throw parseFailure(compoundRootStreams(await readBytes(file)).has(ENCRYPTION_INFO_STREAM) ? 'encrypted' : 'legacyDoc');
    }

    /**
     * Path of the main document part: the target of the package's officeDocument
     * relationship in _rels/.rels, which Word usually but not always names word/document.xml
     * @returns {Promise<string>}
     */
    async findMainDocument() {
        const relsFile = this.zip.file('_rels/.rels');
        if (relsFile) {
            const relsDoc = parseXml(await this.readXmlPart(relsFile));
            for (const rel of relsDoc.getElementsByTagName('Relationship')) {
                if (!OFFICE_DOCUMENT_RELATIONSHIPS.includes(rel.getAttribute('Type'))) continue;
                // Targets in _rels/.rels are relative to the package root
                const path = resolvePartName('', rel.getAttribute('Target') || '');
                if (this.zip.file(path)) return path;
            }
        }
        return 'word/document.xml';
    }

    /**
     * Parse the main document part's relationships (see readRelationships)
     */
    async parseRelationships() {
        this.relationships = await this.readRelationships(this.documentPath);
    }

    /**
     * Read a part's relationships, each with the name of the part it targets (path), resolved
     * relative to that part, or null for an external target such as a hyperlink
     * @param {string} partPath - The part's name, e.g. word/header1.xml
     * @returns {Promise<Object<string, {target: string, type: string, path: string|null}>>} Keyed
     *   by relationship id; empty when the part has no relationships
     */
    async readRelationships(partPath) {
        const relationships = {};
        // A part's relationships live in _rels/<part name>.rels next to it
        const relsFile = this.zip.file(partPath.replace(/[^/]*$/, name => `_rels/${name}.rels`));
        if (!relsFile) return relationships;

        const relsXml = await this.readXmlPart(relsFile);
        const relsDoc = parseXml(relsXml);

        for (const rel of relsDoc.getElementsByTagName('Relationship')) {
            const id = rel.getAttribute('Id');
            const target = rel.getAttribute('Target');
            const type = rel.getAttribute('Type');
            const path = rel.getAttribute('TargetMode') === 'External' || !target
                ? null : resolvePartName(partPath, target);
            relationships[id] = { target, type, path };
        }
        return relationships;
    }

    /**
     * Name of the supporting part the main document part relates to with this
     * relationship type, or where Word puts it when the document has no such relationship
     * @param {string} type - The relationship type, after PART_RELATIONSHIP_TYPE
     * @param {string} fallback - The part's usual name
     * @returns {string}
     */
    relatedPart(type, fallback) {
        const rel = Object.values(this.relationships).find(rel => rel.type === PART_RELATIONSHIP_TYPE + type && rel.path);
        return rel ? rel.path : fallback;
    }

    /**
     * Read a part's bytes within the decompression limits: refused ('tooLarge') before
     * decompressing when its declared size passes the per-part limit or what is left of the
     * total for this parse, and stopped early when it produces more than either allows. A
     * part read again counts toward the total once.
     * @param {JSZipObject} file - The part
     * @returns {Promise<Uint8Array>}
     */
    async readPart(file) {
        const counted = this.partBytes.get(file.name) || 0;
        const limit = Math.min(this.limits.part, this.limits.total - (this.bytesRead - counted));
        const declared = declaredSize(file);
        if (declared !== null && declared > limit) throw partReadFailure('tooLarge');
        const bytes = await inflatePart(file, limit, declared);
        this.bytesRead += bytes.length - counted;
        this.partBytes.set(file.name, bytes.length);
        return bytes;
    }

    /**
     * Read an XML part as text, within the decompression limits
     * @param {JSZipObject} file - The part
     * @returns {Promise<string>}
     */
    async readXmlPart(file) {
        return decodeXmlPart(await this.readPart(file));
    }

    /**
     * Decode the images a part shows, keyed by part name: those an a:blip embeds or a VML
     * v:imagedata shows (legacy pictures and OLE object previews). An image part nothing
     * shows (left over from a deleted picture, or in the mc:AlternateContent branch that
     * isn't read) is never decompressed. An image already decoded isn't read again.
     * @param {Array<Element>} [containers] - Where the pictures are: the main document by
     *   default, text-box pictures included
     * @param {Object} [relationships] - The relationships of the part the containers are in
     *   (see readRelationships)
     * @param {boolean} [skipTextBoxes] - Leave out pictures inside a text box (w:txbxContent),
     *   as in a header or footer, whose text boxes aren't converted
     */
    async extractImages(containers = [this.document], relationships = this.relationships, skipTextBoxes = false) {
        const references = containers.flatMap(container => [
            ...Array.from(container.getElementsByTagNameNS(DRAWING_NS, 'blip'))
                .map(blip => [blip, blip.getAttributeNS(RELATIONSHIPS_NS, 'embed')]),
            ...Array.from(container.getElementsByTagNameNS(VML_NS, 'imagedata'))
                .map(imageData => [imageData, imageData.getAttributeNS(RELATIONSHIPS_NS, 'id')])
        ]);
        const embeds = new Set(references
            .filter(([element, id]) => id && this.isInReadBranch(element) && !(skipTextBoxes && this.isInTextBox(element, null)))
            .map(([, id]) => id));
        for (const id of embeds) {
            if (!Object.prototype.hasOwnProperty.call(relationships, id)) continue;
            const { type, path } = relationships[id];
            if (type !== PART_RELATIONSHIP_TYPE + 'image' || !path || this.images[path]) continue;
            const file = this.zip.file(path);
            if (!file) continue;

            const data = await this.readPart(file);
            this.images[path] = arrayBufferToBase64(data);
            this.imageFormats[path] = imageFormat(data);
        }
    }

    /**
     * Record one occurrence of a conversion finding, where it is: at this.location unless the
     * caller gives one. Nothing is recorded while parsing a cell the generator writes empty (see
     * parseTable): its content is never converted. Nor while a note's paragraphs are parsed.
     * @param {string} code - A FINDING_CODES value
     * @param {string} [example] - A short location hint; blank and repeated ones aren't kept as examples
     * @param {Object} [location] - Where it is: { kind: 'paragraph', paragraph, heading? },
     *   { kind: 'cell', table, row, cell, heading? }, { kind: 'header' }, { kind: 'footer' } or
     *   { kind: 'document' }; a note's text adds note: { kind: 'footnote' | 'endnote' }
     */
    addFinding(code, example = '', location = this.location) {
        if (this.inDiscardedCell || this.parsingNote) return;
        const finding = this.findings[code] ||
            (this.findings[code] = { code, count: 0, examples: [], occurrences: [], places: new Set() });
        finding.count++;
        if (example && finding.examples.length < MAX_FINDING_EXAMPLES && !finding.examples.includes(example)) {
            finding.examples.push(example);
        }
        const place = JSON.stringify([example, location]);
        if (finding.occurrences.length < MAX_FINDING_OCCURRENCES && !finding.places.has(place)) {
            finding.places.add(place);
            finding.occurrences.push({ example, location: copyLocation(location) });
        }
    }

    /**
     * The findings recorded during this parse, in FINDING_CODES order
     * @returns {Array<{code: string, count: number, examples: Array<string>,
     *   occurrences: Array<{example: string, location: Object}>}>} Empty when nothing applies.
     *   occurrences lists up to MAX_FINDING_OCCURRENCES places in the order they were found,
     *   each example and location pair once (see addFinding).
     */
    listFindings() {
        return FINDING_CODES.filter(code => this.findings[code]).map(code => ({
            code,
            count: this.findings[code].count,
            examples: [...this.findings[code].examples],
            occurrences: this.findings[code].occurrences.map(({ example, location }) => ({ example, location: copyLocation(location) }))
        }));
    }

    /**
     * Whether an XML part holds any non-whitespace text (w:t, or an equation's m:t), or a
     * picture, text box, chart or shape (see PICTURE_ELEMENTS)
     * @param {string} path - Part name
     * @returns {Promise<boolean>} false for a missing or unreadable part
     */
    async partHasContent(path) {
        const file = path && this.zip.file(path);
        if (!file) return false;
        try {
            const doc = parseXml(await this.readXmlPart(file));
            return PICTURE_ELEMENTS.some(name => doc.getElementsByTagNameNS(WORD_NS, name).length > 0) ||
                [WORD_NS, MATH_NS].some(namespace => Array.from(doc.getElementsByTagNameNS(namespace, 't'))
                    .some(t => (t.textContent || '').trim() !== ''));
        } catch (error) {
            rethrowFatal(error);
            console.warn(`Could not read ${path}:`, error);
            return false;
        }
    }

    /**
     * The sections of the body, in order: the w:sectPr of each section break (a w:sectPr in a
     * body paragraph's w:pPr), then the body's own final w:sectPr, which describes the last
     * section. The last section exists without one, so it is null then.
     * @returns {Array<Element|null>} Empty without a w:body
     */
    bodySections() {
        const body = this.document.getElementsByTagNameNS(WORD_NS, 'body')[0];
        if (!body) return [];
        const breaks = Array.from(body.getElementsByTagNameNS(WORD_NS, 'sectPr')).filter(sectPr => {
            const pPr = sectPr.parentElement;
            return pPr && pPr.namespaceURI === WORD_NS && pPr.localName === 'pPr' &&
                pPr.parentElement && pPr.parentElement.namespaceURI === WORD_NS && pPr.parentElement.localName === 'p';
        });
        return [...breaks, childElements(body, 'sectPr').pop() || null];
    }

    /**
     * The header and footer references of a section
     * @param {Element|null} sectPr - The section's w:sectPr
     * @returns {Array<{kind: 'header'|'footer', type: string, path: string|null}>} type is
     *   default, first or even; path is null when the relationship is missing or external
     */
    headerFooterReferences(sectPr) {
        if (!sectPr) return [];
        return ['header', 'footer'].flatMap(kind => childElements(sectPr, `${kind}Reference`).map(reference => {
            const rel = this.relationships[reference.getAttributeNS(RELATIONSHIPS_NS, 'id')];
            return { kind, type: reference.getAttribute('w:type') || 'default', path: (rel && rel.path) || null };
        }));
    }

    /**
     * Parse the default header and footer of the document's final section, the section the
     * page setup comes from. A section without its own default header or footer uses the
     * previous section's (ECMA-376 Part 1, 17.10.5). Header and footer content that isn't
     * converted is noted once for the document (headersFooters): text or pictures in a
     * first-page, even-page or other section's header or footer, and fields, text boxes,
     * watermarks and tables in the converted ones. Pictures in the converted ones are kept.
     * @returns {Promise<{header: Array<Object>|null, footer: Array<Object>|null}>} Parsed
     *   paragraphs, or null when there is no text or picture to write
     */
    async parseHeadersFooters() {
        const sections = this.bodySections();
        const converted = { header: null, footer: null };
        for (const kind of ['header', 'footer']) {
            for (let i = sections.length - 1; i >= 0; i--) {
                const reference = this.headerFooterReferences(sections[i])
                    .find(ref => ref.kind === kind && ref.type === 'default');
                if (reference) {
                    converted[kind] = reference.path;
                    break;
                }
            }
        }

        const result = { header: null, footer: null };
        let lost = false;
        for (const kind of ['header', 'footer']) {
            if (!converted[kind]) continue;
            // What the part's own paragraphs report is located in it
            this.location = { kind };
            try {
                const { paragraphs, dropped } = await this.parseHeaderFooterPart(converted[kind], kind);
                result[kind] = paragraphs;
                lost = lost || dropped;
            } finally {
                this.location = { kind: 'document' };
                this.partRelationships = null;
            }
        }

        const checked = new Set();
        for (const { kind, path } of sections.flatMap(sectPr => this.headerFooterReferences(sectPr))) {
            if (lost) break;
            if (!path || path === converted[kind] || checked.has(path)) continue;
            checked.add(path);
            lost = await this.partHasContent(path);
        }

        if (lost) this.addFinding('headersFooters', '', { kind: 'document' });
        return result;
    }

    /**
     * Parse the top-level paragraphs of a header or footer part like body paragraphs, pictures
     * included: they resolve through the part's own relationships, and only the ones it shows
     * are decoded. Fields, text boxes, watermarks and tables are left out: a field's displayed
     * result is dropped, not kept as static text.
     * @param {string} path - The part's name
     * @param {'header'|'footer'} kind - Which part it is: its root is w:hdr or w:ftr
     * @returns {Promise<{paragraphs: Array<Object>|null, dropped: boolean}>} paragraphs is null
     *   when the part is missing, unreadable or has no non-whitespace text or picture left;
     *   dropped says whether any content was left out
     */
    async parseHeaderFooterPart(path, kind) {
        const file = this.zip.file(path);
        if (!file) return { paragraphs: null, dropped: false };
        let root;
        try {
            root = parseXml(await this.readXmlPart(file)).documentElement;
        } catch (error) {
            rethrowFatal(error);
            console.warn(`Could not read ${path}:`, error);
            return { paragraphs: null, dropped: false };
        }
        if (!root || root.namespaceURI !== WORD_NS || root.localName !== (kind === 'header' ? 'hdr' : 'ftr')) {
            return { paragraphs: null, dropped: false };
        }

        let dropped = false;
        const blocks = this.contentChildren(root);
        const hasContent = element => PICTURE_ELEMENTS.some(name => element.getElementsByTagNameNS(WORD_NS, name).length > 0) ||
            Array.from(element.getElementsByTagNameNS(WORD_NS, 't')).some(t => (t.textContent || '').trim() !== '');
        if (blocks.some(block => block.localName === 'tbl' && hasContent(block))) dropped = true;
        // Embedded content from another program (w:altChunk) isn't converted
        if (blocks.some(block => block.localName === 'altChunk')) dropped = true;
        const paragraphs = blocks.filter(block => block.localName === 'p');

        // Simple fields, with the result runs inside them, and watermarks
        const remove = elements => {
            for (const element of elements) {
                dropped = true;
                element.remove();
            }
        };
        const isWatermark = element => Array.from(element.getElementsByTagNameNS(VML_NS, '*'))
            .some(shape => WATERMARK_SHAPE_IDS.test(shape.getAttribute('id') || ''));
        for (const para of paragraphs) {
            remove(Array.from(para.getElementsByTagNameNS(WORD_NS, 'fldSimple')));
            for (const name of PICTURE_ELEMENTS) remove(Array.from(para.getElementsByTagNameNS(WORD_NS, name)).filter(isWatermark));
            // A text box's paragraphs aren't parsed (see parseParagraph); a picture beside it in
            // the same drawing still is (see parseDrawing)
            if (para.getElementsByTagNameNS(WORD_NS, 'txbxContent').length > 0) dropped = true;
        }

        // Complex fields: everything from a w:fldChar begin to its matching end, the field
        // instructions and result included, run content in document order. Text in the same
        // run before the begin or after the end is kept. A field can span paragraphs.
        let depth = 0;
        for (const para of paragraphs) {
            for (const run of this.paragraphRuns(para)) {
                for (const child of Array.from(run.children)) {
                    const isFldChar = child.namespaceURI === WORD_NS && child.localName === 'fldChar';
                    if (isFldChar) {
                        const type = child.getAttribute('w:fldCharType');
                        if (type === 'begin') depth++;
                        else if (type === 'end' && depth > 0) depth--;
                        remove([child]);
                    } else if (depth > 0 && !(child.namespaceURI === WORD_NS && child.localName === 'rPr')) {
                        remove([child]);
                    }
                }
            }
        }

        const hasText = paragraphs.some(para => this.paragraphRuns(para)
            .some(run => run.namespaceURI === MATH_NS ? this.mathText(run).trim() !== ''
                : childElements(run, 'sym').length > 0 || childElements(run, 't').some(t => (t.textContent || '').trim() !== '')));
        // The runs the paragraphs show: not deleted or moved away (see paragraphRuns), nor hidden
        const shownRuns = paragraphs.flatMap(para => {
            const styleRun = this.paragraphStyle(childElement(para, 'pPr')).run;
            return this.paragraphRuns(para)
                .filter(run => run.namespaceURI === WORD_NS && !this.parseRunFormatting(run, styleRun).hidden);
        });
        const hasPicture = shownRuns.some(run => PICTURE_ELEMENTS.some(name => Array.from(run.getElementsByTagNameNS(WORD_NS, name))
            .some(element => this.isInReadBranch(element) && !this.isInTextBox(element, run))));
        if (!hasText && !hasPicture) {
            // Nothing is left to parse, but a part emptied by unaccepted deletions still has changes to report
            for (const para of paragraphs) {
                if (this.hasRevisions(para)) this.addFinding('trackedChanges');
            }
            return { paragraphs: null, dropped };
        }

        // The pictures those runs show, through the part's own relationships: a header's r:embed
        // ids name entries of its _rels/header1.xml.rels, not of the document's
        this.partRelationships = await this.readRelationships(path);
        await this.extractImages(shownRuns, this.partRelationships, true);

        // Every paragraph is written, as in a table cell (see paragraphIsWritten), except one
        // Word's Accept All removes (see parseParagraphBlocks). A table, which isn't converted,
        // still separates the paragraphs on either side of it for contextual spacing.
        const parsed = [];
        const sequence = [];
        for (const block of blocks) {
            if (block.localName !== 'p') {
                if (block.localName === 'tbl' || block.localName === 'altChunk') sequence.push(null);
                continue;
            }
            const paragraph = this.parseParagraph(block, this.paragraphRuns(block), true);
            if (paragraphMarkDeleted(block) && nothingLeft(paragraph.runs)) continue;
            if (paragraph.runs.some(run => run.omittedPictures)) dropped = true;
            paragraph.runs = paragraph.runs.filter(run => HEADER_FOOTER_RUN_TYPES.includes(run.type));
            parsed.push(paragraph);
            sequence.push(paragraph);
        }
        this.applyContextualSpacing(sequence);
        // A part whose only text is hidden (w:vanish), or whose pictures are all missing, shows nothing
        if (!parsed.some(paragraph => paragraph.runs.some(showsContent))) return { paragraphs: null, dropped };
        return { paragraphs: parsed, dropped };
    }

    /**
     * Note comments, which are not converted: one occurrence per comment with text in the
     * comments part (word/comments.xml)
     */
    async detectComments() {
        const commentsFile = this.zip.file(this.relatedPart('comments', 'word/comments.xml'));
        if (!commentsFile) return;
        try {
            const commentsDoc = parseXml(await this.readXmlPart(commentsFile));
            for (const comment of commentsDoc.getElementsByTagNameNS(WORD_NS, 'comment')) {
                const hasText = Array.from(comment.getElementsByTagNameNS(WORD_NS, 't'))
                    .some(t => (t.textContent || '').trim() !== '');
                if (hasText) this.addFinding('comments', '', { kind: 'document' });
            }
        } catch (error) {
            rethrowFatal(error);
            console.warn('Could not read comments:', error);
        }
    }

    /**
     * Text of the paragraph nearest to an element, as a finding's location hint: the
     * paragraph holding it, else the closest paragraph before or after it with text
     * @param {Element} element - An element in the main document part
     * @returns {string} '' when no paragraph has text
     */
    nearestParagraphText(element) {
        let para = element;
        while (para && !(para.namespaceURI === WORD_NS && para.localName === 'p')) para = para.parentElement;
        if (!para) return '';
        // The paragraphs of the part the element is in: the main document, or a header or footer
        let paragraphs;
        if (para.ownerDocument === this.document) {
            if (!this.documentParagraphs) {
                this.documentParagraphs = Array.from(this.document.getElementsByTagNameNS(WORD_NS, 'p'));
            }
            paragraphs = this.documentParagraphs;
        } else {
            paragraphs = Array.from(para.ownerDocument.getElementsByTagNameNS(WORD_NS, 'p'));
        }
        const index = paragraphs.indexOf(para);
        for (let distance = 0; distance < paragraphs.length; distance++) {
            for (const i of distance === 0 ? [index] : [index - distance, index + distance]) {
                if (i < 0 || i >= paragraphs.length) continue;
                const text = shortExample(this.paragraphRuns(paragraphs[i])
                    .flatMap(run => childElements(run, 't'))
                    .map(t => t.textContent || '')
                    .join(''));
                if (text) return text;
            }
        }
        return '';
    }

    /**
     * Parse the notes of the footnotes part (word/footnotes.xml) or the endnotes part
     * (word/endnotes.xml) into plain text keyed by w:id. Word's separator notes are left out.
     * The findings of each note's text (its symbols and hidden text) go into noteFindings, and
     * its formatted paragraphs into noteParagraphs.
     * @param {'footnote'|'endnote'} type - The note element's local name
     * @returns {Promise<Object<string, string>>}
     */
    async parseNotes(type) {
        const notes = {};
        const notesFile = this.zip.file(this.relatedPart(`${type}s`, `word/${type}s.xml`));
        if (!notesFile) return notes;

        try {
            const notesDoc = parseXml(await this.readXmlPart(notesFile));
            // The notes' links resolve through the notes part's own relationships. Notes whose
            // relationships can't be read are still converted; only their links go unreported.
            let relationships = {};
            try {
                relationships = await this.readRelationships(notesFile.name);
            } catch (error) {
                rethrowFatal(error);
                console.warn(`Could not read the ${type}s' relationships:`, error);
            }

            for (const note of notesDoc.getElementsByTagNameNS(WORD_NS, type)) {
                const id = note.getAttribute('w:id');
                // Separator and continuation separator notes are lines, not text. Word
                // gives them ids -1 and 0, and marks them with w:type.
                if (id === '0' || id === '-1' || SEPARATOR_NOTE_TYPES.includes(note.getAttribute('w:type'))) continue;

                const findings = [];
                // Prune the note's own XML before both passes. Removing the whole row also
                // keeps its text out of the plain-text fallback and finding examples.
                for (const row of Array.from(note.getElementsByTagNameNS(WORD_NS, 'tr'))) {
                    if (!note.contains(row) || !tableRowDeleted(childElement(row, 'trPr'))) continue;
                    row.remove();
                    findings.push(['trackedChanges', '']);
                }
                notes[id] = this.noteText(note, findings, relationships);
                this.noteFindings[type][id] = findings;
                this.noteParagraphs[type][id] = this.parseNoteParagraphs(note);
            }
        } catch (error) {
            rethrowFatal(error);
            console.warn(`Could not parse ${type}s:`, error);
        }
        return notes;
    }

    /**
     * Plain text of a footnote or endnote: the text of its runs in document order, with one
     * space between paragraphs. A tab (w:tab, w:ptab) or line break (w:br, w:cr) in a run becomes a
     * space, so the words on either side of it stay apart. A symbol (w:sym) becomes its Unicode
     * text (see symbolText), and an equation (m:oMath) its linear text (see mathText). A hidden run
     * (w:vanish, resolved through the paragraph and character styles as in the body) adds nothing.
     * @param {Element} note - The w:footnote or w:endnote element
     * @param {Array<[string, string]>} [findings] - Receives a [code, example] pair for each symbol
     *   (symbols), each equation (equations), each paragraph with hidden content (hiddenText), with
     *   unaccepted changes (trackedChanges, see hasRevisions), with subscript text (subscript), with
     *   double-struck text (doubleStrikethrough), with a border (paragraphBorders, see hasParagraphBorder),
     *   and for each link whose address its text doesn't show (hyperlinks, see unshownLinkAddresses)
     * @param {Object} [relationships] - The notes part's relationships (see readRelationships)
     * @returns {string}
     */
    noteText(note, findings = [], relationships = {}) {
        let text = '';
        for (const para of note.getElementsByTagNameNS(WORD_NS, 'p')) {
            const start = text.length;
            const pPr = childElement(para, 'pPr');
            const style = this.paragraphStyle(pPr);
            const hidden = { found: false, text: '' };
            let subscript = false;
            let doubleStruck = false;
            // Runs, and equations in place of the runs inside them. Deleted and moved-away runs
            // are left out, as in the body (see paragraphRuns).
            const runs = Array.from(para.getElementsByTagName('*')).filter(element =>
                ((element.namespaceURI === WORD_NS && element.localName === 'r') ||
                    (element.namespaceURI === MATH_NS && element.localName === 'oMath')) && !isInEquation(element, para) &&
                !isDeletedRun(element, para));
            for (const run of runs) {
                if (run.namespaceURI === MATH_NS) {
                    // Its hidden runs add their text to hidden
                    const math = this.equationText(run, style.run, hidden);
                    if (math.trim() !== '') {
                        text += math;
                        findings.push(['equations', shortExample(math)]);
                        if (this.parseRunFormatting(run, style.run).doubleStrike) doubleStruck = true;
                    }
                    continue;
                }
                const formatting = this.parseRunFormatting(run, style.run);
                if (formatting.hidden) {
                    const { shows, text: hiddenText } = hiddenRunContent(run);
                    if (shows) {
                        hidden.found = true;
                        hidden.text += hiddenText;
                    }
                    continue;
                }
                const runStart = text.length;
                for (const child of run.getElementsByTagNameNS(WORD_NS, '*')) {
                    if (child.localName === 't') {
                        text += child.textContent || '';
                    } else if (child.localName === 'sym' && child.parentNode === run) {
                        const symbol = symbolText(child, formatting.fontFamily);
                        text += symbol.text;
                        findings.push(['symbols', symbol.example]);
                    } else if (NOTE_WHITESPACE_ELEMENTS.includes(child.localName) && child.parentNode === run) {
                        text += ' ';
                    }
                }
                const shown = text.slice(runStart).trim() !== '';
                if (formatting.vertAlign === 'subscript' && shown) subscript = true;
                if (formatting.doubleStrike && shown) doubleStruck = true;
            }
            if (hidden.found) findings.push(['hiddenText', shortExample(hidden.text)]);
            for (const example of this.unshownLinkAddresses(para, this.paragraphRuns(para), style.run, relationships)) {
                findings.push(['hyperlinks', example]);
            }
            if (this.hasRevisions(para)) findings.push(['trackedChanges', shortExample(text.slice(start).trim())]);
            if (subscript) findings.push(['subscript', shortExample(text.slice(start).trim())]);
            if (doubleStruck) findings.push(['doubleStrikethrough', shortExample(text.slice(start).trim())]);
            if (this.hasParagraphBorder(style.paragraph, this.readParagraphProperties(pPr))) {
                findings.push(['paragraphBorders', shortExample(text.slice(start).trim())]);
            }
            text += ' ';
        }
        return text.trim();
    }

    /**
     * The paragraphs of a footnote or endnote, each a list of the parsed items (see parseRun) it
     * shows: formatted text, tabs and line breaks. Runs are parsed as in the body, through the
     * document defaults and the paragraph and character styles (e.g. FootnoteText), so hidden
     * text, All Caps, symbols and non-breaking hyphens behave as they do there. Word's note mark
     * (w:footnoteRef, w:endnoteRef) adds nothing, and the spaces or the one tab Word puts after it
     * at the start of the first paragraph are dropped. Page breaks, pictures and nested note
     * references are left out. A run whose size nothing sets is NOTE_FONT_SIZE points.
     * Findings are left to noteText, which records them for each reference to the note.
     * @param {Element} note - The w:footnote or w:endnote element
     * @returns {Array<Array<Object>>} One list of items per w:p, in document order
     */
    parseNoteParagraphs(note) {
        const paragraphs = [];
        const parsingNote = this.parsingNote;
        this.parsingNote = true;
        try {
            for (const para of note.getElementsByTagNameNS(WORD_NS, 'p')) {
                const style = this.paragraphStyle(childElement(para, 'pPr'));
                const paragraphRun = 'fontSize' in this.styles.defaults.run ? style.run : { fontSize: NOTE_FONT_SIZE, ...style.run };
                const items = this.paragraphRuns(para).flatMap(run => this.parseRun(run, paragraphRun));
                // A paragraph Word's Accept All removes (see parseParagraphBlocks) is left out
                if (paragraphMarkDeleted(para) && nothingLeft(items)) continue;
                paragraphs.push(items.filter(item => NOTE_RUN_TYPES.includes(item.type)));
            }
        } finally {
            this.parsingNote = parsingNote;
        }

        // The whitespace after the note mark: leading spaces, and at most one tab
        const first = paragraphs[0] || [];
        let tabDropped = false;
        while (first.length > 0) {
            const item = first[0];
            if (item.type === 'tab' && !tabDropped) {
                tabDropped = true;
                first.shift();
            } else if (item.type === 'text' && item.text.startsWith(' ')) {
                const rest = item.text.replace(/^ +/, '');
                if (rest) first[0] = { ...item, text: rest };
                else first.shift();
            } else {
                break;
            }
        }
        return paragraphs;
    }

    /**
     * Number format of the endnote reference marks: w:endnotePr/w:numFmt in the body's
     * final w:sectPr, then in the settings part. Word's default is lower-case Roman
     * numerals; a format the generator can't write falls back to decimal.
     * @returns {Promise<string>} An ENDNOTE_FORMATS value
     */
    async parseEndnoteFormat() {
        const numFmt = parent => {
            const endnotePr = parent && childElement(parent, 'endnotePr');
            const format = endnotePr && childElement(endnotePr, 'numFmt');
            return format && format.hasAttribute('w:val') ? format.getAttribute('w:val') : null;
        };

        const body = this.document.getElementsByTagNameNS(WORD_NS, 'body')[0];
        const format = numFmt(body && childElements(body, 'sectPr').pop()) ?? numFmt(await this.settingsRoot());
        if (format === null) return DEFAULT_ENDNOTE_FORMAT;
        return ENDNOTE_FORMATS.includes(format) ? format : 'decimal';
    }

    /**
     * Whether the converted endnotes end up somewhere other than Word puts them, or numbered
     * differently. The generator writes every endnote at the end of the document, numbered from
     * the start. That changes a document with more than one section that places endnotes at the
     * end of each section (w:endnotePr/w:pos "sectEnd") or restarts their numbering at each
     * section (w:numRestart "eachSect"), and any document that restarts them at each page.
     * The restart is each section's own: w:numRestart is read from its w:sectPr, then from the
     * settings part. The placement is the document's: w:pos is read from the settings part
     * only, because a w:pos in a section's w:endnotePr is ignored (ECMA-376 Part 1, 17.11.22).
     * @returns {Promise<boolean>} false for Word's default: at the end of the document, numbered continuously
     */
    async endnotesAreMoved() {
        const sections = this.bodySections();
        if (sections.length === 0) return false;

        const value = (parent, name) => {
            const endnotePr = parent && childElement(parent, 'endnotePr');
            const property = endnotePr && childElement(endnotePr, name);
            return property && property.hasAttribute('w:val') ? property.getAttribute('w:val') : null;
        };
        const settings = await this.settingsRoot();
        const manySections = sections.length > 1;
        if (manySections && value(settings, 'pos') === 'sectEnd') return true;

        return sections.some(sectPr => {
            const restart = value(sectPr, 'numRestart') ?? value(settings, 'numRestart') ?? 'continuous';
            return restart === 'eachPage' || (manySections && restart === 'eachSect');
        });
    }

    /**
     * Parse list definitions from the numbering part (word/numbering.xml) into
     * this.numbering[numId][ilvl] = { format, lvlText, bullet, picture, start } (see parseLevel)
     */
    async parseNumbering() {
        const numberingFile = this.zip.file(this.relatedPart('numbering', 'word/numbering.xml'));
        if (!numberingFile) return;

        try {
            const numberingDoc = parseXml(await this.readXmlPart(numberingFile));
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

                // A level override replaces the whole abstract level, defaults included. A start
                // override (Word's Restart at or Set Numbering Value) replaces only its start value.
                for (const override of childElements(num, 'lvlOverride')) {
                    const ilvl = override.getAttribute('w:ilvl') || '0';
                    const lvl = childElement(override, 'lvl');
                    if (lvl) levels[ilvl] = this.parseLevel(lvl);
                    const start = integerValue(childElement(override, 'startOverride'));
                    if (start !== null && levels[ilvl]) levels[ilvl] = { ...levels[ilvl], start };
                }

                this.numbering[num.getAttribute('w:numId')] = levels;
            }
        } catch (error) {
            rethrowFatal(error);
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
     * Whether the settings part (word/settings.xml) puts the gutter at the top of the
     * page (w:gutterAtTop)
     * @returns {Promise<boolean>}
     */
    async isGutterAtTop() {
        const settings = await this.settingsRoot();
        const gutterAtTop = settings && childElement(settings, 'gutterAtTop');
        return gutterAtTop ? this.isOn(gutterAtTop) : false;
    }

    /**
     * The w:settings element of the settings part (word/settings.xml)
     * @returns {Promise<Element|null>} null without a settings part, or for a malformed one
     */
    async settingsRoot() {
        const settingsFile = this.zip.file(this.relatedPart('settings', 'word/settings.xml'));
        if (!settingsFile) return null;

        const root = parseXml(await this.readXmlPart(settingsFile)).documentElement;
        return root && root.namespaceURI === WORD_NS && root.localName === 'settings' ? root : null;
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
     * Read the number format, level text, bullet and start value of a numbering level
     * @param {Element} lvl - The w:lvl element
     * @returns {{format?: string, lvlText?: string, bullet?: string, picture?: boolean, start?: number}}
     *   Only the values the level sets. A bullet level (numFmt bullet) also has bullet, the character
     *   its level text draws (see bulletCharacter), and picture, true for a picture bullet
     *   (w:lvlPicBulletId). start is the w:start value.
     */
    parseLevel(lvl) {
        const level = {};
        const numFmt = childElement(lvl, 'numFmt');
        if (numFmt && numFmt.hasAttribute('w:val')) level.format = numFmt.getAttribute('w:val');
        const lvlText = childElement(lvl, 'lvlText');
        if (lvlText && lvlText.hasAttribute('w:val')) level.lvlText = lvlText.getAttribute('w:val');
        if (level.format === 'bullet') {
            // The level's font: w:ascii, else w:hAnsi
            const rPr = childElement(lvl, 'rPr');
            const rFonts = rPr && childElement(rPr, 'rFonts');
            const font = rFonts ? rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi') || '' : '';
            level.bullet = bulletCharacter(level.lvlText || '', font);
            if (childElement(lvl, 'lvlPicBulletId')) level.picture = true;
        }
        const start = integerValue(childElement(lvl, 'start'));
        if (start !== null) level.start = start;
        return level;
    }

    /**
     * Style tables for a document without styles.xml: nothing set anywhere
     */
    emptyStyles() {
        return {
            defaults: { run: {}, paragraph: {} },
            definitions: { paragraph: new Map(), character: new Map(), table: new Map() },
            defaultParagraph: null,
            defaultTable: null,
            resolved: { paragraph: new Map(), character: new Map(), table: new Map() },
            // Whether a styles part was read: without one, table borders aren't known (see tableBorders)
            loaded: false
        };
    }

    /**
     * Read the theme part's (word/theme/theme1.xml) major and minor Latin typefaces
     * (a:fontScheme a:majorFont / a:minorFont a:latin@typeface), which a w:rFonts
     * w:asciiTheme names. Without a theme part, or with a malformed one, both stay null,
     * and a w:asciiTheme that names them is ignored (see readRunProperties).
     */
    async parseThemeFonts() {
        const themeFile = this.zip.file(this.relatedPart('theme', 'word/theme/theme1.xml'));
        if (!themeFile) return;

        try {
            const root = parseXml(await this.readXmlPart(themeFile)).documentElement;
            if (!root || root.namespaceURI !== DRAWING_NS || root.localName !== 'theme') {
                throw new Error('The theme part is not a well-formed a:theme part');
            }
            const themeElements = childElement(root, 'themeElements', DRAWING_NS);
            const fontScheme = themeElements && childElement(themeElements, 'fontScheme', DRAWING_NS);
            for (const slot of ['major', 'minor']) {
                const font = fontScheme && childElement(fontScheme, `${slot}Font`, DRAWING_NS);
                const latin = font && childElement(font, 'latin', DRAWING_NS);
                const typeface = latin && (latin.getAttribute('typeface') || '').trim();
                this.themeFonts[slot] = typeface || null;
            }
        } catch (error) {
            rethrowFatal(error);
            this.themeFonts = { major: null, minor: null };
            console.warn('Could not parse the theme:', error);
        }
    }

    /**
     * Parse the styles part (word/styles.xml): the document defaults (w:docDefaults),
     * the paragraph and character styles, each with the properties it sets itself, and
     * the table styles' borders. w:basedOn chains are resolved on first use (see
     * resolveStyle and tableBorders).
     */
    async parseStyles() {
        const stylesFile = this.zip.file(this.relatedPart('styles', 'word/styles.xml'));
        if (!stylesFile) return;

        try {
            const root = parseXml(await this.readXmlPart(stylesFile)).documentElement;
            // A malformed part parses to a parsererror document instead of throwing
            if (root.namespaceURI !== WORD_NS || root.localName !== 'styles') {
                throw new Error('The styles part is not a well-formed w:styles part');
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
                const type = style.getAttribute('w:type') || 'paragraph'; // An omitted w:type means paragraph (17.7.4.17)
                const definitions = this.styles.definitions[type];
                const id = style.getAttribute('w:styleId');
                if (!(type === 'paragraph' || type === 'character' || type === 'table') || !id || definitions.has(id)) continue;

                const basedOn = childElement(style, 'basedOn');
                if (type === 'table') {
                    const tblPr = childElement(style, 'tblPr');
                    const tcPr = childElement(style, 'tcPr');
                    definitions.set(id, {
                        basedOn: basedOn ? basedOn.getAttribute('w:val') : null,
                        borders: this.readBorders(tblPr && childElement(tblPr, 'tblBorders'), TABLE_EDGES),
                        cellBorders: this.readBorders(tcPr && childElement(tcPr, 'tcBorders'), CELL_EDGES)
                    });
                    if (this.styles.defaultTable === null && this.isOn(style, 'w:default', false)) {
                        this.styles.defaultTable = id;
                    }
                    continue;
                }

                definitions.set(id, {
                    basedOn: basedOn ? basedOn.getAttribute('w:val') : null,
                    run: this.readRunProperties(childElement(style, 'rPr')),
                    paragraph: type === 'paragraph' ? this.readParagraphProperties(childElement(style, 'pPr')) : {},
                    numPr: type === 'paragraph' ? this.readNumberingProperties(childElement(style, 'pPr')) : {}
                });

                if (type === 'paragraph' && this.styles.defaultParagraph === null && this.isOn(style, 'w:default', false)) {
                    this.styles.defaultParagraph = id;
                }
            }
            this.styles.loaded = true;
        } catch (error) {
            rethrowFatal(error);
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
     * @returns {{run: Object, paragraph: Object, numPr: Object}} Only the properties set somewhere
     */
    resolveStyle(type, styleId) {
        const resolved = this.styles.resolved[type];
        if (!resolved.has(styleId)) {
            const definitions = this.styles.definitions[type];
            const chain = [];
            for (let id = styleId; definitions.has(id) && !chain.includes(definitions.get(id)); id = definitions.get(id).basedOn) {
                chain.unshift(definitions.get(id));
            }

            const paragraph = Object.assign({}, ...chain.map(style => style.paragraph));
            // Tab stops add to and clear the inherited ones instead of replacing them, so the
            // chain's tab operations are all kept, outermost style first
            const tabs = chain.flatMap(style => style.paragraph.tabs || []);
            if (tabs.length > 0) paragraph.tabs = tabs;
            else delete paragraph.tabs;
            // Border edges inherit one by one, not as a whole w:pBdr
            const borderEdges = Object.assign({}, ...chain.map(style => style.paragraph.borderEdges));
            if (Object.keys(borderEdges).length > 0) paragraph.borderEdges = borderEdges;
            else delete paragraph.borderEdges;

            resolved.set(styleId, {
                run: Object.assign({}, ...chain.map(style => style.run)),
                paragraph,
                numPr: Object.assign({}, ...chain.map(style => style.numPr))
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
        return this.resolveStyle('paragraph', this.paragraphStyleId(pPr));
    }

    /**
     * The id of a paragraph's paragraph style: its w:pStyle when styles.xml defines it, else the
     * default paragraph style's
     * @param {Element|null} pPr - The paragraph's w:pPr
     * @returns {string|null} null when there is no style at all
     */
    paragraphStyleId(pPr) {
        const pStyle = pPr && childElement(pPr, 'pStyle');
        const id = pStyle && pStyle.getAttribute('w:val');
        return this.styles.definitions.paragraph.has(id) ? id : this.styles.defaultParagraph;
    }

    /**
     * Read the character properties a w:rPr sets, from a run, a style or the document defaults
     * @param {Element|null} rPr - The w:rPr element
     * @returns {Object} Only the properties it sets; an explicit "off" is kept as false, and
     *   color "auto", highlight "none" and shading "nil" or "auto" as null, so they override an earlier layer
     */
    readRunProperties(rPr) {
        const props = {};
        if (!rPr) return props;

        const onOff = [['b', 'bold'], ['i', 'italic'], ['strike', 'strike'], ['dstrike', 'doubleStrike'],
            ['vanish', 'hidden'], ['caps', 'caps'], ['smallCaps', 'smallCaps']];
        for (const [name, key] of onOff) {
            const el = childElement(rPr, name);
            if (el) props[key] = this.isOn(el);
        }

        // The language of the run's text (w:lang w:val, e.g. "tr-TR"), which decides how All Caps upper-cases it
        const lang = childElement(rPr, 'lang');
        if (lang && lang.getAttribute('w:val')) props.lang = lang.getAttribute('w:val');

        const u = childElement(rPr, 'u');
        if (u && u.hasAttribute('w:val')) props.underline = u.getAttribute('w:val') !== 'none';

        // Superscript or subscript (w:vertAlign); an explicit baseline cancels an inherited one
        const vertAlign = childElement(rPr, 'vertAlign');
        const vertAlignVal = vertAlign && vertAlign.getAttribute('w:val');
        if (VERTICAL_ALIGNMENTS.includes(vertAlignVal)) props.vertAlign = vertAlignVal;

        // The Latin font: a theme font (w:asciiTheme) wins over w:ascii in the same w:rFonts
        // (ECMA-376 Part 1, 17.3.2.26). One the theme doesn't give falls back to w:ascii, else
        // leaves the font to the layer underneath.
        const rFonts = childElement(rPr, 'rFonts');
        if (rFonts) {
            const slot = THEME_FONT_SLOTS[rFonts.getAttribute('w:asciiTheme')];
            const themeFont = slot && this.themeFonts[slot];
            if (themeFont) props.fontFamily = themeFont;
            else if (rFonts.getAttribute('w:ascii')) props.fontFamily = rFonts.getAttribute('w:ascii');
        }

        // Font size is in half-points
        const sz = childElement(rPr, 'sz');
        const halfPoints = sz ? parseInt(sz.getAttribute('w:val')) : NaN;
        if (halfPoints > 0) props.fontSize = halfPoints / 2;

        const color = childElement(rPr, 'color');
        const colorVal = color && color.getAttribute('w:val');
        if (colorVal) props.color = colorVal === 'auto' ? null : '#' + colorVal;

        // Highlighting (w:highlight): "none" is null, so it overrides an inherited highlight;
        // a value outside Word's palette is ignored
        const highlight = childElement(rPr, 'highlight');
        const highlightVal = highlight && highlight.getAttribute('w:val');
        if (highlightVal === 'none') props.highlight = null;
        else if (Object.hasOwn(HIGHLIGHT_COLORS, highlightVal)) props.highlight = '#' + HIGHLIGHT_COLORS[highlightVal];

        // Run shading (w:shd): a solid pattern is drawn entirely in its pattern color (w:color),
        // any other pattern is approximated by its background fill (w:fill). "nil", or an "auto"
        // or missing color, is null; a color that isn't 6 hex digits is ignored.
        const shd = childElement(rPr, 'shd');
        if (shd) {
            const pattern = shd.getAttribute('w:val');
            const solidColor = pattern === 'solid' ? shd.getAttribute('w:color') : '';
            const fill = solidColor && solidColor !== 'auto' ? solidColor : shd.getAttribute('w:fill');
            if (pattern === 'nil' || !fill || fill === 'auto') props.shading = null;
            else if (/^[0-9A-Fa-f]{6}$/.test(fill)) props.shading = '#' + fill.toUpperCase();
        }

        return props;
    }

    /**
     * Read the alignment, indents and spacing a w:pPr sets, from a paragraph, a style or the document defaults
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
            // A hanging indent is a negative first-line indent (see hangingIndent). It replaces an
            // inherited first-line indent, and wins over w:firstLine on the same w:ind (ECMA-376 17.3.1.12).
            if (ind.hasAttribute('w:hanging')) {
                props.firstLineIndent = -this.twipsToPoints(ind.getAttribute('w:hanging')) || 0;
            } else if (ind.hasAttribute('w:firstLine')) {
                props.firstLineIndent = this.twipsToPoints(ind.getAttribute('w:firstLine'));
            }
        }

        const spacing = childElement(pPr, 'spacing');
        if (spacing) {
            // Space before and after, in points to a tenth; an unreadable or negative value is ignored
            const space = attribute => {
                const points = this.measureToPoints(spacing.getAttribute(attribute));
                return points !== null && points >= 0 ? Math.round(points * 10) / 10 : null;
            };
            const before = space('w:before');
            if (before !== null) props.spaceBefore = before;
            const after = space('w:after');
            if (after !== null) props.spaceAfter = after;

            // The line height and its rule inherit separately; lineSpacing combines them after the merge
            const line = Number(spacing.getAttribute('w:line'));
            if (spacing.hasAttribute('w:line') && Number.isFinite(line) && line > 0) props.line = line;
            if (spacing.getAttribute('w:lineRule')) props.lineRule = spacing.getAttribute('w:lineRule');
        }

        const tabs = this.readTabs(childElement(pPr, 'tabs'));
        if (tabs.length > 0) props.tabs = tabs;

        // The outline level, which makes a paragraph a heading (see MAX_HEADING_OUTLINE_LEVEL)
        const outlineLvl = childElement(pPr, 'outlineLvl');
        const outlineLevel = outlineLvl ? parseInt(outlineLvl.getAttribute('w:val'), 10) : NaN;
        if (Number.isInteger(outlineLevel)) props.outlineLevel = outlineLevel;

        // Page break before (w:pageBreakBefore), which starts a body paragraph on a new page (see bodyPageBreakBefore)
        const pageBreakBefore = childElement(pPr, 'pageBreakBefore');
        if (pageBreakBefore) props.pageBreakBefore = this.isOn(pageBreakBefore);

        // Contextual spacing (w:contextualSpacing), which drops the space between paragraphs of
        // the same style (see applyContextualSpacing)
        const contextualSpacing = childElement(pPr, 'contextualSpacing');
        if (contextualSpacing) props.contextualSpacing = this.isOn(contextualSpacing);

        // Paragraph borders (w:pBdr), which the UDF can't draw: whether each edge it sets draws a
        // line. Edges inherit one by one (see resolveStyle and hasParagraphBorder).
        const pBdr = childElement(pPr, 'pBdr');
        if (pBdr) {
            const edges = {};
            for (const edge of PARAGRAPH_BORDER_EDGES) {
                const element = childElement(pBdr, edge);
                if (element) edges[edge] = !INVISIBLE_BORDER_STYLES.includes(element.getAttribute('w:val') || 'nil');
            }
            if (Object.keys(edges).length > 0) props.borderEdges = edges;
        }

        return props;
    }

    /**
     * Read the tab stop operations of a w:tabs: each w:tab adds a stop, or with
     * w:val="clear" removes the inherited stop at its position (see resolveTabStops).
     * A bar stop is kept as { position, bar: true }, which converts to no stop, and a stop
     * whose leader has no UDF code is marked lostLeader, for lostTabStops. Stops with an
     * unreadable or negative position, or an unknown w:val, are left out.
     * @param {Element|null} tabs - The w:tabs element
     * @returns {Array<{position: number, clear?: true, bar?: true, alignment?: number, leader?: number,
     *   lostLeader?: true}>} Positions in points, not rounded: a stop is replaced or cleared
     *   only at its own position (see tabStopKey), and two stops a twip apart round to the same tenth
     */
    readTabs(tabs) {
        const operations = [];
        for (const tab of tabs ? childElements(tabs, 'tab') : []) {
            const position = this.measureToPoints(tab.getAttribute('w:pos'));
            if (position === null || !Number.isFinite(position) || position < 0) continue;

            const val = tab.getAttribute('w:val');
            if (val === 'clear') {
                operations.push({ position, clear: true });
            } else if (val === 'bar') {
                operations.push({ position, bar: true });
            } else if (Object.hasOwn(TAB_ALIGNMENTS, val)) {
                const leader = tab.getAttribute('w:leader');
                const operation = {
                    position,
                    alignment: TAB_ALIGNMENTS[val],
                    leader: Object.hasOwn(TAB_LEADERS, leader) ? TAB_LEADERS[leader] : 0
                };
                if (LOST_TAB_LEADERS.includes(leader)) operation.lostLeader = true;
                operations.push(operation);
            }
        }
        return operations;
    }

    /**
     * Apply tab stop operations in order: a stop replaces any stop at its position,
     * and a clear removes it. Bar stops replace inherited stops but position no text
     * and are not converted. Positions are compared by tabStopKey, whatever unit they were
     * written in, and rounded to a tenth of a point only in the result. Stops that are left
     * less than a tenth apart can round to the same position, which TabSet gives no meaning
     * to: only the first of them in position order, the one a tab reaches first, is kept.
     * @param {Array<Object>} operations - See readTabs, outermost layer first
     * @returns {Array<{position: number, alignment: number, leader: number}>} Sorted by position,
     *   positions in points, to a tenth, each position once
     */
    resolveTabStops(operations) {
        const stops = new Map();
        for (const { position, clear, bar, alignment, leader } of operations) {
            const key = tabStopKey(position);
            if (clear || bar) stops.delete(key);
            else stops.set(key, { position, alignment, leader });
        }

        const resolved = [];
        let lastTenths = null;
        for (const stop of [...stops.values()].sort((a, b) => a.position - b.position)) {
            const tenths = Math.round(stop.position * 10);
            if (tenths === lastTenths) continue;
            resolved.push({ ...stop, position: tenths / 10 });
            lastTenths = tenths;
        }
        return resolved;
    }

    /**
     * Whether the tab stops a paragraph ends up with include one the conversion loses: a bar
     * stop, or a stop with an underscore or heavy leader. Layers apply as in Word: a later stop
     * at the same position replaces it, and a clear removes it. Positions are compared as in
     * resolveTabStops, by tabStopKey.
     * @param {Array<Object>} operations - See readTabs, outermost layer first
     * @returns {boolean}
     */
    hasLostTabStops(operations) {
        const lost = new Set();
        for (const { position, clear, bar, lostLeader } of operations) {
            const key = tabStopKey(position);
            if (!clear && (bar || lostLeader)) lost.add(key);
            else lost.delete(key);
        }
        return lost.size > 0;
    }

    /**
     * Read the list a w:pPr puts its paragraph in, from a paragraph or a paragraph style.
     * w:numId and w:ilvl inherit separately, so each is kept only if set.
     * @param {Element|null} pPr - The w:pPr element
     * @returns {{numId?: string, level?: number}} Only the values it sets
     */
    readNumberingProperties(pPr) {
        const props = {};
        const numPr = pPr && childElement(pPr, 'numPr');
        if (!numPr) return props;

        const numId = childElement(numPr, 'numId');
        if (numId && numId.hasAttribute('w:val')) props.numId = numId.getAttribute('w:val');

        // A w:ilvl without a readable level is level 0, as before
        const ilvl = childElement(numPr, 'ilvl');
        if (ilvl) {
            const level = parseInt(ilvl.getAttribute('w:val'), 10);
            props.level = level >= 0 ? level : 0;
        }

        return props;
    }

    /**
     * The line spacing multiplier of resolved paragraph properties. w:line counts 240ths
     * of a line when the rule is auto (or absent). Exact and at-least line heights have
     * no UDF equivalent, so they give single spacing.
     * @param {{line?: number, lineRule?: string}} props
     * @returns {number} 1 for single spacing
     */
    lineSpacing({ line, lineRule = 'auto' }) {
        return lineRule === 'auto' && line ? line / 240 : 1;
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

        try {
            return this.parseBodyBlocks(body);
        } finally {
            this.location = { kind: 'document' };
            // A change to the final section's properties (margins, orientation, headers), which
            // are converted as they now are, belongs to the whole document
            const sectPr = childElement(body, 'sectPr');
            if (sectPr && childElement(sectPr, 'sectPrChange')) this.addFinding('trackedChanges', '', { kind: 'document' });
        }
    }

    /**
     * Parse the body's own paragraphs and tables, keeping this.location on the one being parsed,
     * so its findings say where they are. Body paragraphs and tables are numbered from 1 in
     * document order, empty paragraphs included; a cell is located by its table, its row and its
     * w:tc position in that row. Text boxes and nested tables aren't counted: their findings take
     * the location of the paragraph that anchors the text box, or of the outermost table's cell. Each
     * location names the last heading before it with an outline level, else the last petition
     * heading before it (see headingKind). Embedded content (w:altChunk), which isn't converted,
     * is numbered as a body paragraph, so its finding names where it starts.
     * @param {Element} body - The w:body element
     * @returns {Array<Object>} Parsed paragraphs and tables in document order (see parseBlocks)
     */
    parseBodyBlocks(body) {
        const elements = [];
        const headingOf = () => {
            const { outline, petition } = this.body.headings;
            const heading = outline !== null ? outline : petition;
            return heading === null ? {} : { heading };
        };

        // Each section break's w:sectPr, and the w:sectPr of the section that starts after it, whose
        // w:type says whether that section starts on a new page
        const children = this.contentChildren(body);
        const finalSection = childElements(body, 'sectPr').pop() || null;
        const breaks = children.map(child => child.localName === 'p' ? sectionBreakOf(child) : null).filter(Boolean);
        const nextSection = new Map(breaks.map((sectPr, i) => [sectPr, i + 1 < breaks.length ? breaks[i + 1] : finalSection]));
        // The body's own paragraphs, for contextual spacing; null where a table or embedded content is.
        // Page breaks and section breaks don't separate paragraphs: Word applies it across them.
        const sequence = [];

        for (const child of children) {
            if (child.localName === 'p') {
                this.location = { kind: 'paragraph', paragraph: ++this.body.paragraphs, ...headingOf() };
                const { paragraph, blocks } = this.parseParagraphBlocks(child);
                if (paragraph) sequence.push(paragraph);
                // Page breaks are block-level pageBreak items between the paragraphs and tables: the
                // generator writes one only when something follows it, and never two in a row
                if (paragraph && this.bodyPageBreakBefore(child)) elements.push({ type: 'pageBreak' });
                elements.push(...blocks);
                const sectPr = sectionBreakOf(child);
                if (paragraph && sectPr) {
                    const start = sectionStart(nextSection.get(sectPr));
                    if (PAGE_SECTION_STARTS.includes(start)) elements.push({ type: 'pageBreak' });
                    if (PARITY_SECTION_STARTS.includes(start) || this.sectionLayoutDiffers(sectPr, finalSection) ||
                        hasColumns(sectPr)) {
                        this.addFinding('sectionLayout', shortExample(paragraphText(paragraph)));
                    }
                }
                // A heading names the places after it, not its own. A paragraph left out isn't one.
                const kind = paragraph && this.headingKind(child, paragraph);
                if (kind) this.body.headings[kind] = shortExample(paragraphText(paragraph));
            } else if (child.localName === 'tbl') {
                sequence.push(null);
                const table = ++this.body.tables;
                const heading = headingOf();
                const parsed = this.parseTable(child, (row, cell) => {
                    this.location = { kind: 'cell', table, row, cell, ...heading };
                });
                if (parsed) elements.push(parsed);
            } else if (child.localName === 'altChunk') {
                sequence.push(null);
                // Embedded content takes a body paragraph's place: Word shows it where its first paragraph would be
                this.location = { kind: 'paragraph', paragraph: ++this.body.paragraphs, ...headingOf() };
                this.reportEmbeddedChunk(child);
            }
        }

        // The final section's columns: the <pageFormat> is written from its page setup, but not its columns
        if (finalSection && hasColumns(finalSection)) this.addFinding('sectionLayout', '', { kind: 'document' });

        this.applyContextualSpacing(sequence);
        return elements;
    }

    /**
     * Whether a body paragraph starts on a new page (w:pageBreakBefore): its own w:pPr, else its
     * paragraph style chain, else the document defaults. An explicit off turns an inherited one off.
     * @param {Element} para - The w:p element
     * @returns {boolean}
     */
    bodyPageBreakBefore(para) {
        const pPr = childElement(para, 'pPr');
        const layers = [this.readParagraphProperties(pPr), this.paragraphStyle(pPr).paragraph, this.styles.defaults.paragraph];
        const layer = layers.find(props => props.pageBreakBefore !== undefined);
        return layer ? layer.pageBreakBefore : false;
    }

    /**
     * Whether a section's page setup differs from the final section's, which the <pageFormat> is
     * written from: its orientation, its page width or height (w:pgSz), or its top, right, bottom
     * or left margin (w:pgMar). Values are compared as parsed, in twips, so one written in other
     * units isn't a difference. Only values both sections set are compared: one either leaves out, or that can't
     * be read, isn't a known difference. The orientation is compared when both have a w:pgSz.
     * @param {Element} sectPr - The section's w:sectPr
     * @param {Element|null} finalSection - The body's final w:sectPr
     * @returns {boolean}
     */
    sectionLayoutDiffers(sectPr, finalSection) {
        const setup = section => {
            const pgSz = section && childElement(section, 'pgSz');
            const pgMar = section && childElement(section, 'pgMar');
            // To the twip, so the same length written in other units compares equal
            const read = (el, attribute) => {
                const points = el ? this.measureToPoints(el.getAttribute(attribute)) : null;
                return points === null ? null : Math.round(points * 20);
            };
            const [width, height] = PAGE_SIZE_ATTRIBUTES.map(attribute => read(pgSz, attribute));
            const landscape = pgSz ? pgSz.getAttribute('w:orient') === 'landscape' ||
                (width !== null && height !== null && width > height) : null;
            return [landscape, width, height, ...PAGE_MARGIN_ATTRIBUTES.map(attribute => read(pgMar, attribute))];
        };
        const own = setup(sectPr);
        const final = setup(finalSection);
        return own.some((value, i) => value !== null && final[i] !== null && value !== final[i]);
    }

    /**
     * Whether a body paragraph is a heading, and of which kind: 'outline' when its outline level,
     * from its own w:pPr else its paragraph style chain, is one of Word's heading levels (built-in
     * Heading 1-9 styles set one), else 'petition' when it reads as a petition heading (see
     * isPetitionHeading). A paragraph with no text is never one.
     * @param {Element} para - The w:p element
     * @param {Object} paragraph - Its parsed paragraph
     * @returns {'outline'|'petition'|null}
     */
    headingKind(para, paragraph) {
        if (paragraphText(paragraph).trim() === '') return null;
        const pPr = childElement(para, 'pPr');
        const own = this.readParagraphProperties(pPr).outlineLevel;
        const level = own !== undefined ? own : this.paragraphStyle(pPr).paragraph.outlineLevel;
        if (level >= 0 && level <= MAX_HEADING_OUTLINE_LEVEL) return 'outline';
        return isPetitionHeading(paragraph) ? 'petition' : null;
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
     * Parse the block-level content (paragraphs and tables) of a body, cell or text box. Embedded
     * content (w:altChunk) is reported (see reportEmbeddedChunk), not converted.
     * @param {Element} container - The containing element
     * @param {boolean} [inCell] - Whether the content ends up in a table cell
     * @returns {Array<Object>} Parsed paragraphs and tables in document order
     */
    parseBlocks(container, inCell = false) {
        const elements = [];
        // The container's own paragraphs, for contextual spacing; null where a table or embedded content is
        const sequence = [];

        for (const child of this.contentChildren(container)) {
            const localName = child.localName;

            if (localName === 'p') {
                const { paragraph, blocks } = this.parseParagraphBlocks(child, inCell);
                if (paragraph) sequence.push(paragraph);
                elements.push(...blocks);
            } else if (localName === 'tbl') {
                sequence.push(null);
                const table = this.parseTable(child);
                if (table) elements.push(table);
            } else if (localName === 'altChunk') {
                sequence.push(null);
                this.reportEmbeddedChunk(child);
            }
        }

        this.applyContextualSpacing(sequence);
        return elements;
    }

    /**
     * Apply contextual spacing (w:contextualSpacing, ECMA-376 17.3.1.9) to a sequence of
     * paragraphs: when a paragraph that has it is next to one of the same paragraph style, its
     * space on that side is dropped. Each paragraph's own setting decides only its own spacing:
     * the first of two loses its space after, the second its space before. Paragraphs are
     * adjacent when they follow each other in the sequence; a paragraph left out isn't in it (see
     * parseParagraphBlocks), and text boxes anchored in a paragraph are sequences of their own.
     * @param {Array<Object|null>} sequence - Parsed paragraphs in document order, with null where
     *   a table (or other block) stands between two of them
     */
    applyContextualSpacing(sequence) {
        for (let i = 1; i < sequence.length; i++) {
            const previous = sequence[i - 1];
            const next = sequence[i];
            const before = previous && this.contextualSpacing.get(previous);
            const after = next && this.contextualSpacing.get(next);
            if (!before || !after || before.style !== after.style) continue;
            if (before.on && previous.spaceAfter) previous.spaceAfter = 0;
            if (after.on && next.spaceBefore) next.spaceBefore = 0;
        }
    }

    /**
     * Parse a paragraph, then the blocks of the text boxes anchored in it, which follow it. A
     * paragraph whose mark is deleted and which has nothing left once its deleted runs are dropped
     * is left out, as Word's Accept All removes it; its text boxes are still parsed.
     * @param {Element} para - The w:p element
     * @param {boolean} [inCell] - Whether the content ends up in a table cell
     * @returns {{paragraph: Object|null, blocks: Array<Object>}} The parsed paragraph, null when it
     *   is left out, and the blocks: the paragraph first, then the text boxes' paragraphs and tables
     */
    parseParagraphBlocks(para, inCell = false) {
        const runs = this.paragraphRuns(para);
        let paragraph = this.parseParagraph(para, runs, inCell);
        if (paragraphMarkDeleted(para) && nothingLeft(paragraph.runs)) paragraph = null;
        const blocks = paragraph ? [paragraph] : [];
        for (const run of runs) {
            for (const textBox of this.findTextBoxes(run)) {
                blocks.push(...this.parseBlocks(textBox, inCell));
            }
        }
        return { paragraph, blocks };
    }

    /**
     * Whether a paragraph holds Track Changes revision marks (see PARAGRAPH_REVISION_ELEMENTS):
     * in its runs, its inline wrappers or its properties. The paragraphs of a text box anchored
     * in it are left to their own parse, and only the read branch of mc:AlternateContent counts.
     * @param {Element} element - The w:p element, or an element inside it
     * @returns {boolean}
     */
    hasRevisions(element) {
        for (const child of element.children) {
            if (child.namespaceURI === MARKUP_COMPATIBILITY_NS && child.localName === 'AlternateContent') {
                const branch = this.alternateContentBranch(child);
                if (branch && this.hasRevisions(branch)) return true;
            } else if (child.namespaceURI === WORD_NS && child.localName === 'txbxContent') {
                continue;
            } else if (child.namespaceURI === WORD_NS && PARAGRAPH_REVISION_ELEMENTS.includes(child.localName)) {
                return true;
            } else if (this.hasRevisions(child)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Collect the runs of a paragraph, including runs inside inline wrappers
     * (hyperlinks, insertions, content controls, smart tags, fields, ...).
     * Deleted and moved-away text is skipped. An equation (m:oMath, also each one in an
     * m:oMathPara) is collected whole in place of its runs: parseRun writes it as linear text.
     * @param {Element} container - The paragraph or an inline wrapper
     * @param {Array<Element>} [runs] - Accumulator
     * @returns {Array<Element>} w:r and m:oMath elements in document order
     */
    paragraphRuns(container, runs = []) {
        const skipped = ['pPr', 'rPr', 'sdtPr', 'sdtEndPr', 'customXmlPr', 'smartTagPr', 'del', 'moveFrom'];

        for (const child of container.children) {
            if (child.namespaceURI === MARKUP_COMPATIBILITY_NS) {
                if (child.localName === 'AlternateContent') {
                    const branch = this.alternateContentBranch(child);
                    if (branch) this.paragraphRuns(branch, runs);
                }
            } else if (child.namespaceURI === MATH_NS && child.localName === 'oMath') {
                runs.push(child);
            } else if (child.namespaceURI !== WORD_NS) {
                // e.g. an m:oMathPara, which holds equations
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
     * Whether an element is outside mc:AlternateContent, or only in branches that
     * alternateContentBranch picks to read
     * @param {Element} element - An element of a parsed part
     * @returns {boolean}
     */
    isInReadBranch(element) {
        for (let el = element.parentElement; el; el = el.parentElement) {
            if (el.namespaceURI === MARKUP_COMPATIBILITY_NS && (el.localName === 'Choice' || el.localName === 'Fallback') &&
                el.parentElement && this.alternateContentBranch(el.parentElement) !== el) {
                return false;
            }
        }
        return true;
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
     * @param {boolean} [inCell] - Whether the paragraph ends up in a table cell
     * @returns {Object} Parsed paragraph
     */
    parseParagraph(para, runs = this.paragraphRuns(para), inCell = false) {
        const paragraph = {
            type: 'paragraph',
            alignment: 'left',
            runs: [],
            numbering: null
        };

        // Alignment, indents and spacing: document defaults, the paragraph style chain, then the paragraph's own w:pPr
        const pPr = childElement(para, 'pPr');
        const style = this.paragraphStyle(pPr);
        const own = this.readParagraphProperties(pPr);
        Object.assign(paragraph, this.styles.defaults.paragraph, style.paragraph, own);
        paragraph.lineSpacing = this.lineSpacing(paragraph);

        // Contextual spacing is settled against the neighbouring paragraphs once the sequence the
        // paragraph is in has been parsed (see applyContextualSpacing)
        this.contextualSpacing.set(paragraph, { style: this.paragraphStyleId(pPr), on: paragraph.contextualSpacing === true });
        delete paragraph.contextualSpacing;

        // The outline level only locates findings (see headingKind); it isn’t written. Page break
        // before is written as a page break ahead of a body paragraph (see bodyPageBreakBefore).
        // Borders can't be written; they are reported below.
        delete paragraph.outlineLevel;
        delete paragraph.pageBreakBefore;
        delete paragraph.borderEdges;

        // Tab stops: each layer adds to or clears the stops of the layers under it
        delete paragraph.tabs;
        const tabOperations = [this.styles.defaults.paragraph, style.paragraph, own].flatMap(layer => layer.tabs || []);
        const tabStops = this.resolveTabStops(tabOperations);
        if (tabStops.length > 0) paragraph.tabStops = tabStops;

        // Numbering: the paragraph style chain, then the paragraph's own w:numPr, each value
        // on its own. numId 0 switches numbering off, and a missing w:ilvl means level 0.
        const { numId, level = 0 } = { ...style.numPr, ...this.readNumberingProperties(pPr) };
        if (numId !== undefined && numId !== '0') {
            paragraph.numbering = { level, numId };
            // The list definition, if numbering.xml has one for this level
            const definition = (this.numbering[numId] || {})[level];
            if (definition) {
                // A level without w:numFmt is decimal
                paragraph.numbering.format = definition.format || 'decimal';
                paragraph.numbering.lvlText = definition.lvlText || '';
                for (const key of ['bullet', 'picture', 'start']) {
                    if (definition[key] !== undefined) paragraph.numbering[key] = definition[key];
                }
            }
        }

        this.hangingIndent(paragraph);

        // Get runs (text with formatting). Hidden runs add nothing; their text is collected for the finding.
        const hidden = { found: false, text: '' };
        const doubleStruck = { found: false };
        for (const run of runs) {
            paragraph.runs.push(...this.parseRun(run, style.run, hidden, doubleStruck));
        }
        if (hidden.found) this.addFinding('hiddenText', shortExample(hidden.text));

        // Links whose address the text doesn't show: the UDF has no links, so the address is lost.
        // A header's or footer's links resolve through that part's relationships.
        for (const example of this.unshownLinkAddresses(para, runs, style.run, this.partRelationships || this.relationships)) {
            this.addFinding('hyperlinks', example);
        }

        // Unaccepted changes: the paragraph is converted as if they were accepted
        if (this.hasRevisions(para)) this.addFinding('trackedChanges', shortExample(paragraphText(paragraph)));

        // Subscript text, written on the baseline (see parseRun)
        if (paragraph.runs.some(run => run.subscript && run.text.trim() !== '')) this.addFinding('subscript', shortExample(paragraphText(paragraph)));

        // Double-struck text, written with a single strikethrough (see parseRun)
        if (doubleStruck.found) this.addFinding('doubleStrikethrough', shortExample(paragraphText(paragraph)));

        // Line heights, tab stops and borders the conversion loses, with the paragraph's text as the hint.
        // A paragraph the generator doesn't write as one has no line height, stops or borders to lose.
        if (paragraphIsWritten(paragraph, inCell)) {
            // A list marker the UDF list attributes approximate (see listMarkerApproximated)
            if (paragraph.numbering && listMarkerApproximated(paragraph.numbering)) {
                this.addFinding('listMarkers', shortExample(paragraphText(paragraph)));
            }
            if (this.hasParagraphBorder(style.paragraph, own)) {
                this.addFinding('paragraphBorders', shortExample(paragraphText(paragraph)));
            }
            if (LOST_LINE_RULES.includes(paragraph.lineRule)) {
                this.addFinding('lineSpacing', shortExample(paragraphText(paragraph)));
            }
            if (this.hasLostTabStops(tabOperations)) {
                this.addFinding('tabStops', shortExample(paragraphText(paragraph)));
            }
        }

        return paragraph;
    }

    /**
     * Settle a paragraph's hanging indent (a negative firstLineIndent, see readParagraphProperties).
     * A list paragraph keeps no first-line indent, as before: its ListLevel drives the editor's
     * list indentation, and how that combines with a negative FirstLineIndent is unverified.
     * Any other paragraph keeps it, clamped so that the first line starts no further left than
     * the margin, and gets the tab stop Word implies at the left indent: a left stop without a
     * leader, unless one of its own stops lies past the first line's start and up to the left indent.
     * List paragraphs are those with numbering whose level is a list (format other than none).
     * @param {Object} paragraph - The parsed paragraph, with its numbering and tabStops resolved
     */
    hangingIndent(paragraph) {
        if (!(paragraph.firstLineIndent < 0)) return;
        const numbering = paragraph.numbering;
        if (numbering && numbering.format !== 'none') {
            paragraph.firstLineIndent = 0;
            return;
        }

        const leftIndent = Math.max(paragraph.leftIndent || 0, 0);
        paragraph.firstLineIndent = Math.max(paragraph.firstLineIndent, -leftIndent) || 0;
        if (paragraph.firstLineIndent === 0) return;

        // Indents are whole points (see twipsToPoints), so a stop is compared at the point it rounds to:
        // the w:pos="2268" stop of a w:left="2268" paragraph (113.4 against 113) is at the left indent
        const firstLineStart = leftIndent + paragraph.firstLineIndent;
        const stops = paragraph.tabStops || [];
        const between = stop => Math.round(stop.position) > firstLineStart && Math.round(stop.position) <= leftIndent;
        if (stops.some(between)) return;
        paragraph.tabStops = [...stops, { position: leftIndent, alignment: 0, leader: 0 }]
            .sort((a, b) => a.position - b.position);
    }

    /**
     * The links a paragraph shows whose address its text doesn't show (see linkShowsAddress), as
     * hyperlinks examples: a w:hyperlink whose r:id names an external relationship of the part it is
     * in, and a HYPERLINK field, complex (w:fldChar begin, its w:instrText, separate, the result, end)
     * or simple (w:fldSimple), a field nested in another field's result included. A link's text is
     * what its runs show: hidden runs add nothing, and deleted and moved-away runs aren't among the
     * runs (see paragraphRuns). Internal links (w:anchor only, HYPERLINK \l), links with no text and a
     * w:hyperlink whose relationship is missing aren't listed. Only reads: nothing is recorded.
     * @param {Element} para - The w:p element
     * @param {Array<Element>} runs - Its runs (see paragraphRuns)
     * @param {Object} styleRun - Run properties from the paragraph's style (see resolveStyle)
     * @param {Object} relationships - The relationships of the part the paragraph is in (see readRelationships)
     * @returns {Array<string>} One example per link, in the order the links start: the link's text,
     *   then → and its address without the scheme, e.g. Yargıtay kararı → example.org/karar
     */
    unshownLinkAddresses(para, runs, styleRun, relationships) {
        const links = [];
        const newLink = address => {
            const link = { address, text: '' };
            links.push(link);
            return link;
        };
        // The link of each w:hyperlink and w:fldSimple met, null for one that isn't an external link
        const wrapperLinks = new Map();
        const wrapperLink = element => {
            if (!wrapperLinks.has(element)) {
                let address = '';
                if (element.localName === 'hyperlink') {
                    const id = element.getAttributeNS(RELATIONSHIPS_NS, 'id');
                    const relationship = id && Object.prototype.hasOwnProperty.call(relationships, id) ? relationships[id] : null;
                    // An external target has no part name (see readRelationships)
                    if (relationship && relationship.target && relationship.path === null) address = relationship.target.trim();
                } else {
                    address = hyperlinkFieldAddress(element.getAttribute('w:instr') || '');
                }
                wrapperLinks.set(element, address ? newLink(address) : null);
            }
            return wrapperLinks.get(element);
        };
        // The complex fields open at this point, outermost first
        const fields = [];

        for (const run of runs) {
            if (run.namespaceURI !== WORD_NS) continue;
            // The links of the w:hyperlink and w:fldSimple elements around the run, outermost first
            const around = [];
            for (let element = run.parentElement; element && element !== para; element = element.parentElement) {
                if (element.namespaceURI === WORD_NS && (element.localName === 'hyperlink' || element.localName === 'fldSimple')) {
                    around.unshift(element);
                }
            }
            const wrappers = around.map(wrapperLink).filter(Boolean);
            const hidden = this.parseRunFormatting(run, styleRun).hidden;

            for (const child of run.children) {
                if (child.namespaceURI !== WORD_NS) continue;
                const field = fields[fields.length - 1];
                if (child.localName === 'fldChar') {
                    const type = child.getAttribute('w:fldCharType');
                    if (type === 'begin') {
                        fields.push({ instruction: '', inResult: false, link: null });
                    } else if (type === 'separate' && field && !field.inResult) {
                        field.inResult = true;
                        const address = hyperlinkFieldAddress(field.instruction);
                        if (address) field.link = newLink(address);
                    } else if (type === 'end') {
                        fields.pop();
                    }
                } else if (child.localName === 'instrText') {
                    if (field && !field.inResult) field.instruction += child.textContent || '';
                } else if (!hidden && (child.localName === 't' || NOTE_WHITESPACE_ELEMENTS.includes(child.localName))) {
                    const text = child.localName === 't' ? child.textContent || '' : ' ';
                    for (const link of wrappers) link.text += text;
                    // A field's result shows the results of the fields nested in it
                    for (const open of fields) {
                        if (open.link) open.link.text += text;
                    }
                }
            }
        }

        return links
            .filter(({ text, address }) => text.trim() !== '' && !linkShowsAddress(text, address))
            .map(({ text, address }) => shortExample(`${text.trim()} → ${address.replace(LINK_SCHEME, '')}`));
    }

    /**
     * Whether a paragraph draws a border line (w:pBdr), which the UDF can't write: an edge whose
     * line style isn't nil or none, from the document defaults, the paragraph style chain or the
     * paragraph's own w:pPr. Each edge inherits on its own, so an edge set to none hides only that edge.
     * @param {Object} styleParagraph - The paragraph style's paragraph properties (see resolveStyle)
     * @param {Object} own - The paragraph's own properties (see readParagraphProperties)
     * @returns {boolean}
     */
    hasParagraphBorder(styleParagraph, own) {
        const edges = { ...this.styles.defaults.paragraph.borderEdges, ...styleParagraph.borderEdges, ...own.borderEdges };
        return Object.values(edges).some(Boolean);
    }

    /**
     * Parse a run element (text with formatting)
     * @param {Element} run - The run element
     * @param {Object} [paragraphRun] - Run properties from the paragraph's style (see resolveStyle)
     * @param {{found: boolean, text: string}} [hidden] - Set when the run is hidden (w:vanish) and
     *   holds content: found becomes true and the run's text is appended to text
     * @param {{found: boolean}} [doubleStruck] - found becomes true when the run is double-struck
     *   (w:dstrike) and shows text other than whitespace
     * @returns {Array<Object>} Parsed items in document order (text, tab, break, pageBreak, image,
     *   footnoteRef, endnoteRef). Empty for a hidden run: Word neither shows nor prints it. Text of
     *   a superscript run has superscript: true; text of a subscript run has subscript: true and is
     *   written on the baseline, as the UDF has no subscript. Text of a double-struck run (w:dstrike)
     *   has strike: true, as the UDF has no double strikethrough.
     */
    parseRun(run, paragraphRun = {}, hidden = null, doubleStruck = null) {
        if (run.namespaceURI === MATH_NS) return this.parseEquation(run, paragraphRun, hidden, doubleStruck);
        const items = [];
        // hidden, caps, smallCaps and lang decide how the run's text is written, vertAlign whether
        // it is superscript or subscript, and doubleStrike whether it is struck; they aren't text attributes
        const { hidden: isHidden, caps, smallCaps, lang, vertAlign, doubleStrike, ...formatting } =
            this.parseRunFormatting(run, paragraphRun);
        if (vertAlign === 'superscript') formatting.superscript = true;
        else if (vertAlign === 'subscript') formatting.subscript = true;
        if (doubleStrike) formatting.strike = true;

        if (isHidden) {
            const { shows, text } = hiddenRunContent(run);
            if (hidden && shows) {
                hidden.found = true;
                hidden.text += text;
            }
            return items;
        }

        // All Caps and Small Caps show the text in capitals (the UDF has no small capitals). Word
        // upper-cases by the run's language, so Turkish text, or text with no language set, keeps
        // its dotted capital İ.
        const turkish = !lang || lang.toLowerCase().startsWith('tr');
        const shown = text => !(caps || smallCaps) ? text
            : turkish ? text.toLocaleUpperCase('tr-TR') : text.toUpperCase();
        // asIs keeps the text as it is: a symbol is a fixed glyph, which All Caps doesn't change
        const addText = (text, asIs = false) => {
            if (!text) return;
            const written = asIs ? text : shown(text);
            const last = items[items.length - 1];
            if (last && last.type === 'text') {
                last.text += written;
            } else {
                items.push({ ...formatting, text: written });
            }
        };

        for (const child of run.children) {
            const localName = child.localName;

            if (localName === 't') {
                addText(child.textContent || '');
            } else if (localName === 'noBreakHyphen') {
                // A plain hyphen: whether the UYAP editor's fonts show U+2011 is unknown
                addText('-');
            } else if (localName === 'sym') {
                // Reported even when it maps: the UYAP editor's fonts may draw it differently
                const symbol = symbolText(child, formatting.fontFamily);
                this.addFinding('symbols', symbol.example);
                addText(symbol.text, true);
            } else if (localName === 'tab' || localName === 'ptab') {
                // A positional tab (w:ptab) aligns to a margin Word computes; it is written as a plain tab
                items.push({ type: 'tab' });
            } else if (localName === 'br') {
                const breakType = child.getAttribute('w:type');
                items.push({ type: breakType === 'page' ? 'pageBreak' : 'break' });
            } else if (localName === 'cr') {
                items.push({ type: 'break' });
            } else if (this.parsingNote) {
                // In a note, pictures and nested note references are left out (see parseNoteParagraphs)
                continue;
            } else if (localName === 'footnoteReference' || localName === 'endnoteReference') {
                const kind = localName === 'footnoteReference' ? 'footnote' : 'endnote';
                const notes = kind === 'footnote' ? this.footnotes : this.endnotes;
                const id = child.getAttribute('w:id');
                // A note with no text is still a note; a reference to a missing one adds nothing
                if (id && Object.prototype.hasOwnProperty.call(notes, id)) {
                    // The formatted paragraphs go with a note that has text; an empty one keeps its placeholder
                    const paragraphs = notes[id] ? this.noteParagraphs[kind][id] : null;
                    items.push({
                        type: `${kind}Ref`,
                        id: id,
                        content: notes[id] || `[Empty ${kind}]`,
                        paragraphs: paragraphs || null
                    });
                    // Located at the reference, and marked as being in the note's text
                    const inNote = { ...this.location, note: { kind } };
                    // Footnotes always leave their page; endnotes only when Word places or numbers
                    // them per section, which is a setting of the whole document
                    if (kind === 'footnote') this.addFinding('notesMoved', shortExample(notes[id]), inNote);
                    else if (this.endnotesMoved) this.addFinding('endnotesMoved', shortExample(notes[id]), { kind: 'document' });
                    // The note's text is written once per reference, and its findings with it
                    for (const [code, example] of this.noteFindings[kind][id] || []) this.addFinding(code, example, inNote);
                }
            } else if (localName !== 'rPr') {
                items.push(...this.parsePictures(child));
            }
        }

        if (doubleStrike && doubleStruck && items.some(item => item.type === 'text' && item.text.trim() !== '')) {
            doubleStruck.found = true;
        }
        return items;
    }

    /**
     * Parse an equation (m:oMath) into a text item of its linear text (see mathText), reported as
     * an equations finding. The text takes the paragraph's run formatting, as a run without its
     * own w:rPr would; All Caps doesn't change it, as x and X are different variables. Its hidden
     * runs are left out (see mathText). Each equation after the first in an m:oMathPara starts a
     * new line, as Word shows it.
     * @param {Element} oMath - The m:oMath element
     * @param {Object} [paragraphRun] - Run properties from the paragraph's style (see resolveStyle)
     * @param {{found: boolean, text: string}} [hidden] - See parseRun; set by the equation's hidden runs
     * @param {{found: boolean}} [doubleStruck] - See parseRun; set when the equation is double-struck
     * @returns {Array<Object>} Empty for an equation with no text, or one all hidden
     */
    parseEquation(oMath, paragraphRun = {}, hidden = null, doubleStruck = null) {
        const text = this.equationText(oMath, paragraphRun, hidden);
        if (text.trim() === '') return [];
        // hidden, caps, smallCaps, lang and vertAlign don't apply to the equation's text as a whole;
        // double strikethrough is written as a strikethrough, as for a run
        const { hidden: isHidden, caps, smallCaps, lang, vertAlign, doubleStrike, ...formatting } =
            this.parseRunFormatting(oMath, paragraphRun);
        if (doubleStrike) {
            formatting.strike = true;
            if (doubleStruck) doubleStruck.found = true;
        }

        this.addFinding('equations', shortExample(text));
        const items = [];
        const parent = oMath.parentElement;
        if (parent && parent.namespaceURI === MATH_NS && parent.localName === 'oMathPara' &&
            childElements(parent, 'oMath', MATH_NS)[0] !== oMath) {
            items.push({ type: 'break' });
        }
        items.push({ ...formatting, text });
        return items;
    }

    /**
     * Linear text of an equation (see mathText). An equation whose runs with text are all hidden
     * has none: the operators and brackets of its structures (a fraction's /, a script's ^) would
     * be all that is left.
     * @param {Element} oMath - The m:oMath element
     * @param {Object} [paragraphRun] - Run properties from the paragraph's style (see resolveStyle)
     * @param {{found: boolean, text: string}} [hidden] - Set when a hidden run has text, which is appended
     * @returns {string}
     */
    equationText(oMath, paragraphRun = {}, hidden = null) {
        const hiddenRuns = { found: false, text: '' };
        const visible = { found: false };
        const text = this.mathText(oMath, { paragraphRun, hidden: hiddenRuns, visible });
        if (hiddenRuns.found && hidden) {
            hidden.found = true;
            hidden.text += hiddenRuns.text;
        }
        return hiddenRuns.found && !visible.found ? '' : text;
    }

    /**
     * Linear text of an Office Math element, close to Word's own linear format: a fraction
     * (m:f) is num/den, scripts are base^sup, base_sub and base_sub^sup, a radical (m:rad)
     * √(x) or d√(x) with a degree that isn't hidden, a delimiter (m:d) its parts between its
     * begin and end characters, joined by its separator, an n-ary operator (m:nary) the
     * operator, its limits as _sub and ^sup, a space and its body, and a function (m:func) its
     * name, a space and its argument. A fraction without a bar is binom(n, k), and a matrix
     * (m:m) keeps its rows and cells as matrix([a, b], [c, d]). A fraction part, base, script,
     * degree or limit is wrapped in parentheses unless it is simple (see mathOperand).
     * Other structures (accents, bars, boxes, equation arrays, limits, ...) are read in document order.
     * A hidden run (m:r or w:r whose w:rPr, character style or paragraph style resolves w:vanish
     * on, see parseRunFormatting) adds nothing.
     * @param {Element} element - An m:oMath, or an element inside one
     * @param {{paragraphRun?: Object, hidden?: {found: boolean, text: string}|null,
     *   visible?: {found: boolean}|null}} [context] - paragraphRun: run properties from the paragraph's
     *   style; hidden: set when a hidden run has text, which is appended to hidden.text; visible:
     *   set when a run that isn't hidden has text other than whitespace
     * @returns {string}
     */
    mathText(element, context = {}) {
        const ns = element.namespaceURI;
        const name = element.localName;
        const children = el => this.mathChildrenText(el, context);
        if (ns === MARKUP_COMPATIBILITY_NS) {
            const branch = name === 'AlternateContent' ? this.alternateContentBranch(element) : null;
            return branch ? children(branch) : '';
        }
        if ((ns === MATH_NS || ns === WORD_NS) && name === 'r') {
            const isHidden = this.parseRunFormatting(element, context.paragraphRun || {}).hidden;
            const text = this.mathChildrenText(element, { ...context, hidden: null, visible: null });
            if (isHidden) {
                if (context.hidden && text !== '') {
                    context.hidden.found = true;
                    context.hidden.text += text;
                }
                return '';
            }
            if (context.visible && text.trim() !== '') context.visible.found = true;
            return text;
        }
        if (ns === WORD_NS) {
            if (name === 't') return element.textContent || '';
            if (name === 'sym') return symbolText(element).text;
            if (MATH_SKIPPED_WORD_ELEMENTS.includes(name)) return '';
            return children(element);
        }
        if (ns !== MATH_NS) return children(element);
        // Properties (m:rPr, m:fPr, m:ctrlPr, ...) show no text
        if (name.endsWith('Pr')) return '';

        const part = partName => {
            const el = childElement(element, partName, MATH_NS);
            return el ? children(el) : '';
        };
        const operand = partName => this.mathOperand(childElement(element, partName, MATH_NS), context);
        const hidden = (properties, flag) => {
            const pr = childElement(element, properties, MATH_NS);
            return mathIsOn(pr && childElement(pr, flag, MATH_NS));
        };

        switch (name) {
            case 't':
                return element.textContent || '';
            case 'f': {
                const properties = childElement(element, 'fPr', MATH_NS);
                const type = properties && childElement(properties, 'type', MATH_NS);
                if (type && type.getAttributeNS(MATH_NS, 'val') === 'noBar') {
                    return `binom(${part('num')}, ${part('den')})`;
                }
                return `${operand('num')}/${operand('den')}`;
            }
            case 'sSup':
                return `${operand('e')}^${operand('sup')}`;
            case 'sSub':
                return `${operand('e')}_${operand('sub')}`;
            case 'sSubSup':
                return `${operand('e')}_${operand('sub')}^${operand('sup')}`;
            case 'rad': {
                const degree = hidden('radPr', 'degHide') ? '' : operand('deg');
                return `${degree}√(${part('e')})`;
            }
            case 'd': {
                const begin = mathCharacter(element, 'dPr', 'begChr', '(');
                const end = mathCharacter(element, 'dPr', 'endChr', ')');
                const separator = mathCharacter(element, 'dPr', 'sepChr', '|');
                const parts = childElements(element, 'e', MATH_NS).map(children);
                return `${begin}${parts.join(separator)}${end}`;
            }
            case 'nary': {
                const sub = hidden('naryPr', 'subHide') ? '' : operand('sub');
                const sup = hidden('naryPr', 'supHide') ? '' : operand('sup');
                const body = part('e');
                return `${mathCharacter(element, 'naryPr', 'chr', '∫')}${sub ? `_${sub}` : ''}${sup ? `^${sup}` : ''}` +
                    (body ? ` ${body}` : '');
            }
            case 'func':
                return `${part('fName')} ${part('e')}`;
            case 'm': {
                // Rows and cells can sit in the selected branch of mc:AlternateContent.
                const matrixParts = (parent, name) => Array.from(parent.children).flatMap(child => {
                    if (child.namespaceURI === MATH_NS && child.localName === name) return [child];
                    if (child.namespaceURI !== MARKUP_COMPATIBILITY_NS || child.localName !== 'AlternateContent') return [];
                    const branch = this.alternateContentBranch(child);
                    return branch ? matrixParts(branch, name) : [];
                });
                const rows = matrixParts(element, 'mr').map(row =>
                    `[${matrixParts(row, 'e').map(children).join(', ')}]`);
                return `matrix(${rows.join(', ')})`;
            }
            default:
                return children(element);
        }
    }

    /**
     * Linear text of an element's children, in document order (see mathText)
     * @param {Element} element
     * @param {Object} [context] - See mathText
     * @returns {string}
     */
    mathChildrenText(element, context = {}) {
        return Array.from(element.children).map(child => this.mathText(child, context)).join('');
    }

    /**
     * Linear text of a fraction part, base, script, degree or limit, wrapped in parentheses
     * unless it is empty, a single run of letters or digits or a decimal number (A/365, but
     * (a+b)/2 and x^(n+1)), or already one delimiter (m:d) with both a begin and an end character
     * @param {Element|null} part - e.g. m:num, m:sup
     * @param {Object} [context] - See mathText
     * @returns {string} '' for a missing part
     */
    mathOperand(part, context = {}) {
        if (!part) return '';
        const text = this.mathChildrenText(part, context);
        if (text === '' || SIMPLE_MATH_OPERAND.test(text)) return text;
        const shown = Array.from(part.children).filter(child => !(child.namespaceURI === MATH_NS && child.localName.endsWith('Pr')));
        const bracketed = d => d.namespaceURI === MATH_NS && d.localName === 'd' &&
            mathCharacter(d, 'dPr', 'begChr', '(') !== '' && mathCharacter(d, 'dPr', 'endChr', ')') !== '';
        if (shown.length === 1 && bracketed(shown[0])) return text;
        return `(${text})`;
    }

    /**
     * Report embedded content from another program (w:altChunk: HTML, RTF, MHT or another
     * .docx), which isn't converted, as an embeddedContent finding at this.location. The example
     * is the file name of the part it embeds.
     * @param {Element} altChunk - The w:altChunk element
     */
    reportEmbeddedChunk(altChunk) {
        this.embeddedChunks++;
        const id = altChunk.getAttributeNS(RELATIONSHIPS_NS, 'id');
        const relationship = id && Object.prototype.hasOwnProperty.call(this.relationships, id) ? this.relationships[id] : null;
        const target = relationship ? (relationship.path || relationship.target || '') : '';
        this.addFinding('embeddedContent', shortExample(target.split('/').pop()));
    }

    /**
     * Parse a run child that may show a picture: w:drawing, a VML w:pict or w:object, or
     * mc:AlternateContent, of which only the branch alternateContentBranch picks is read, so
     * a picture written both as DrawingML and as a VML fallback appears once
     * @param {Element} element - A run child
     * @returns {Array<Object>} Parsed images (see parseDrawing), in document order
     */
    parsePictures(element) {
        if (element.namespaceURI === MARKUP_COMPATIBILITY_NS && element.localName === 'AlternateContent') {
            const branch = this.alternateContentBranch(element);
            return branch ? Array.from(branch.children).flatMap(child => this.parsePictures(child)) : [];
        }
        if (element.namespaceURI === WORD_NS && element.localName === 'drawing') {
            return [this.parseDrawing(element)];
        }
        if (element.namespaceURI === WORD_NS && (element.localName === 'pict' || element.localName === 'object')) {
            const image = this.parseVmlPicture(element);
            return image ? [image] : [];
        }
        // Any other element: a drawing somewhere inside it
        const drawing = element.getElementsByTagNameNS(WORD_NS, 'drawing')[0];
        return drawing ? [this.parseDrawing(drawing)] : [];
    }

    /**
     * Resolve a run's character formatting in Word's order: built-in defaults, document
     * defaults, the paragraph style chain, the character style chain (w:rStyle), then
     * the run's own w:rPr. Later layers override earlier ones, except that the on/off
     * properties in TOGGLE_PROPERTIES (bold, italic, strike, hidden, caps and small caps)
     * set by the two style levels toggle: on in both styles is off, and off
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
        // Word draws highlighting over shading, so the highlight wins; each resolves on its own first
        const background = formatting.highlight || formatting.shading;
        delete formatting.highlight;
        delete formatting.shading;
        if (background) formatting.background = background;

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
            width: DEFAULT_IMAGE_SIZE,
            height: DEFAULT_IMAGE_SIZE,
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
        // text box's own paragraphs, which are parsed separately; one in an
        // mc:AlternateContent branch that isn't read isn't shown (or extracted).
        const blips = Array.from(drawing.getElementsByTagNameNS(DRAWING_NS, 'blip'))
            .filter(blip => !this.isInTextBox(blip, drawing) && this.isInReadBranch(blip));
        if (blips.length > 0) {
            this.embedPicture(image, blips[0].getAttributeNS(RELATIONSHIPS_NS, 'embed'), drawing);
            if (blips.length > 1) image.omittedPictures = true;
        } else if (drawing.getElementsByTagNameNS(WORD_NS, 'txbxContent').length === 0) {
            // A chart, SmartArt diagram or shape with no picture and no text box to keep
            this.addFinding('unconvertedDrawing', this.nearestParagraphText(drawing));
        }

        return image;
    }

    /**
     * Parse a VML picture (w:pict) or OLE object (w:object): the first v:imagedata outside a
     * text box is embedded, sized by its shape's style. An OLE object is shown by that preview
     * picture, as in Word; one written with a DrawingML w:drawing instead is parsed as a drawing.
     * @param {Element} element - The w:pict or w:object element
     * @returns {Object|null} Parsed image (see parseDrawing), or null when it shows no picture
     */
    parseVmlPicture(element) {
        const imageData = Array.from(element.getElementsByTagNameNS(VML_NS, 'imagedata'))
            .find(el => el.getAttributeNS(RELATIONSHIPS_NS, 'id') && !this.isInTextBox(el, element) && this.isInReadBranch(el));
        if (imageData) {
            const image = { type: 'image', ...vmlShapeSize(imageData.parentElement), data: null };
            this.embedPicture(image, imageData.getAttributeNS(RELATIONSHIPS_NS, 'id'), element);
            return image;
        }

        const drawing = Array.from(element.getElementsByTagNameNS(WORD_NS, 'drawing'))
            .find(el => !this.isInTextBox(el, element) && this.isInReadBranch(el));
        if (drawing) return this.parseDrawing(drawing);

        // A shape, line or object with no picture; a text box's paragraphs are kept on their own
        if (element.getElementsByTagNameNS(WORD_NS, 'txbxContent').length === 0) {
            this.addFinding('unconvertedDrawing', this.nearestParagraphText(element));
        }
        return null;
    }

    /**
     * Set a parsed image's data from the picture a relationship points at. A picture that
     * isn't PNG or JPEG is embedded as-is, though UDF documents only those, and reported as
     * unsupportedImage; one with no relationship or media part is reported as missingImage
     * and keeps no data, so the generator skips it.
     * @param {Object} image - The parsed image
     * @param {string} relationshipId - r:embed of an a:blip, or r:id of a v:imagedata
     * @param {Element} element - The picture's element, for the location hint
     */
    embedPicture(image, relationshipId, element) {
        // A header's or footer's pictures resolve through that part's relationships
        const relationships = this.partRelationships || this.relationships;
        const relationship = relationshipId && Object.prototype.hasOwnProperty.call(relationships, relationshipId)
            ? relationships[relationshipId] : null;
        const path = relationship && relationship.path;
        if (path && this.images[path]) {
            image.data = this.images[path];
            const format = this.imageFormats[path];
            if (!SUPPORTED_IMAGE_FORMATS.includes(format)) this.addFinding('unsupportedImage', format);
        } else {
            this.addFinding('missingImage', this.nearestParagraphText(element));
        }
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
     * Parse a table element, at this.tableDepth + 1 levels of tables deep. A table within
     * MAX_TABLE_DEPTH levels, nested in a cell or not, is written as a table: the generator
     * writes a cell that continues the vertical merge above it empty, so nothing in such a cell
     * is recorded as a finding. A table nested deeper is flattened into its cell's paragraphs
     * instead (see parseCellBlocks), those cells included, and its borders aren't written.
     * @param {Element} tbl - The table element
     * @param {function(number, number): void} [enterCell] - Called with a cell's 1-based row and
     *   w:tc position in that row before its content is parsed (see parseBodyBlocks)
     * @returns {Object} Parsed table; each cell has its content as blocks (paragraphs and
     *   written nested tables, in document order) and as paragraphs (see flattenToParagraphs)
     */
    parseTable(tbl, enterCell = null) {
        const depth = this.tableDepth + 1;
        const flattened = depth > MAX_TABLE_DEPTH;
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

        const edges = this.tableBorders(childElement(tbl, 'tblPr'));

        // Get rows and cells (only this table's own, not those of nested tables). The cells'
        // properties come first: they decide which cells continue a vertical merge. A row marked
        // deleted is left out, as Word's Accept All removes it, before borders and merges are
        // resolved; rows keep Word's numbers (deleted rows counted) for locating findings.
        const cellElements = [];
        const rowNumbers = [];
        const changedRows = [];
        // The numbers of the deleted rows before each kept row, and after the last one
        const deletedRows = [[]];
        let rowNumber = 0;
        for (const tr of this.contentChildren(tbl)) {
            if (tr.localName !== 'tr') continue;
            rowNumber++;
            const trPr = childElement(tr, 'trPr');
            if (tableRowDeleted(trPr)) {
                deletedRows[deletedRows.length - 1].push(rowNumber);
                continue;
            }
            const tcs = this.contentChildren(tr).filter(tc => tc.localName === 'tc');
            cellElements.push(tcs);
            rowNumbers.push(rowNumber);
            changedRows.push(Boolean(trPr) && ROW_REVISION_ELEMENTS.some(name => childElement(trPr, name)));
            deletedRows.push([]);
            table.rows.push({ type: 'row', cells: tcs.map(tc => this.parseTableCell(tc)) });
        }
        // A deleted row is reported at its first cell, in document order with the kept rows' findings
        const reportDeleted = rows => rows.forEach(number => {
            if (enterCell) enterCell(number, 1);
            this.addFinding('trackedChanges');
        });
        // Every row deleted: Word's Accept All removes the table
        if (rowNumber > 0 && table.rows.length === 0) {
            reportDeleted(deletedRows[0]);
            return null;
        }

        this.resolveCellBorders(table, edges, cellElements);
        if (!table.rows.some(row => row.cells.some(cell => Object.values(cell.borders).some(Boolean)))) {
            table.border = 'borderNone';
        }

        const merges = verticalMerges(table);
        this.tableDepth = depth;
        const firstText = paragraphs => paragraphs.map(paragraphText).find(text => text.trim() !== '') || '';
        try {
            table.rows.forEach((row, i) => {
                reportDeleted(deletedRows[i]);
                row.cells.forEach((cell, j) => {
                    const discarded = !flattened && merges[i][j].joinsAbove;
                    if (enterCell) enterCell(rowNumbers[i], j + 1);
                    // A cell of a nested table may be in a cell written empty: that stays so
                    const wasDiscarded = this.inDiscardedCell;
                    if (discarded) this.inDiscardedCell = true;
                    try {
                        cell.blocks = this.parseCellBlocks(cellElements[i][j]);
                        cell.paragraphs = this.flattenToParagraphs(cell.blocks);
                        // An inserted, deleted or merged cell, or a cell formatting change
                        const tcPr = childElement(cellElements[i][j], 'tcPr');
                        if (tcPr && CELL_REVISION_ELEMENTS.some(name => childElement(tcPr, name))) {
                            this.addFinding('trackedChanges', shortExample(firstText(cell.paragraphs)));
                        }
                    } finally {
                        this.inDiscardedCell = wasDiscarded;
                    }
                    // A flattened table's borders aren't written
                    if (!flattened && table.border !== 'borderNone') this.checkCellBorders(cell, merges[i][j], discarded);
                });
                // An inserted row, or a row formatting change, located at its first cell
                if (changedRows[i]) {
                    if (enterCell) enterCell(rowNumbers[i], 1);
                    this.addFinding('trackedChanges', shortExample(firstText(row.cells.flatMap(cell => cell.paragraphs))));
                }
            });
            reportDeleted(deletedRows[table.rows.length]);
        } finally {
            this.tableDepth = depth - 1;
        }

        return table;
    }

    /**
     * Record a tableBorders finding for a cell whose border attributes can't describe its lines:
     * a line style with no documented equivalent was approximated as solid, or its edges draw
     * different lines and the heaviest was written for the whole cell (see cellBorderChoice)
     * @param {Object} cell - Parsed cell, with its borders and paragraphs
     * @param {{joinsAbove: boolean, joinsBelow: boolean}} merge - The cell's vertical merge
     * @param {boolean} discarded - Whether the generator writes the cell empty (it continues a merge)
     */
    checkCellBorders(cell, merge, discarded) {
        const { approximated, mixed } = cellBorderChoice(cell, merge);
        if (!approximated && !mixed) return;
        const text = discarded ? '' : cell.paragraphs.map(paragraphText).find(t => t.trim() !== '');
        this.addFinding('tableBorders', shortExample(text || ''));
    }

    /**
     * Read which edges a w:tblBorders or w:tcBorders sets, and the line each draws
     * @param {Element|null} borders - The w:tblBorders or w:tcBorders element
     * @param {Object<string,Array<string>>} edgeNames - Edge -> the element names that set it
     * @returns {Object<string,Object|null>} Only the edges it sets (an element without w:val sets
     *   none): null for no line (none, nil), else the line (see readBorderLine)
     */
    readBorders(borders, edgeNames) {
        const set = {};
        if (!borders) return set;
        for (const [edge, names] of Object.entries(edgeNames)) {
            for (const name of names) {
                const border = childElement(borders, name);
                if (border && border.hasAttribute('w:val')) {
                    set[edge] = this.readBorderLine(border);
                }
            }
        }
        return set;
    }

    /**
     * The line a border element (w:top, w:insideH...) draws
     * @param {Element} border - A border element with a w:val
     * @returns {{val: string, size: number, color: string}|null} null when w:val draws no line.
     *   size: w:sz in eighths of a point, 4 (½ pt) when missing or not a positive whole number.
     *   color: w:color as RRGGBB (upper case), or 'auto' when it is auto, missing or not a
     *   hex color. w:themeColor is not resolved.
     */
    readBorderLine(border) {
        const val = border.getAttribute('w:val');
        if (NO_BORDER_VALUES.includes(val)) return null;
        const sz = (border.getAttribute('w:sz') || '').trim();
        const size = /^\d+$/.test(sz) && Number(sz) > 0 ? Number(sz) : DEFAULT_BORDER_LINE.size;
        const color = (border.getAttribute('w:color') || '').trim();
        return { val, size, color: /^[0-9A-Fa-f]{6}$/.test(color) ? color.toUpperCase() : 'auto' };
    }

    /**
     * The line each of a table's six edges draws: its table style (w:tblStyle, else the default
     * table style) with its w:basedOn chain, then the table's own w:tblBorders. The nearest
     * layer that sets an edge wins, and an edge nothing sets has no line. Without a styles
     * part the style layer is unknown, so every edge has Word's default line (single, ½ pt,
     * auto) unless the table sets it.
     * The style's cell borders (its w:tcPr/w:tcBorders, through the same chain) are returned
     * apart: they apply to every cell, over the table's edges. Conditional formatting
     * (w:tblStylePr) is not read.
     * @param {Element|null} tblPr - The table's w:tblPr
     * @returns {{top: ?Object, left: ?Object, bottom: ?Object, right: ?Object, insideH: ?Object,
     *   insideV: ?Object, cells: Object<string,?Object>}} Each edge's line (see readBorderLine), or
     *   null for none. cells: only the cell edges the style sets
     */
    tableBorders(tblPr) {
        let base;
        let cells = {};
        if (this.styles.loaded) {
            const tblStyle = tblPr && childElement(tblPr, 'tblStyle');
            const id = tblStyle && tblStyle.getAttribute('w:val');
            const definitions = this.styles.definitions.table;
            const styleId = definitions.has(id) ? id : this.styles.defaultTable;
            const resolved = this.styles.resolved.table;
            if (!resolved.has(styleId)) {
                const chain = [];
                for (let sid = styleId; definitions.has(sid) && !chain.includes(definitions.get(sid)); sid = definitions.get(sid).basedOn) {
                    chain.unshift(definitions.get(sid));
                }
                resolved.set(styleId, {
                    borders: Object.assign({}, ...chain.map(style => style.borders)),
                    cells: Object.assign({}, ...chain.map(style => style.cellBorders))
                });
            }
            ({ borders: base, cells } = resolved.get(styleId));
        } else {
            base = Object.fromEntries(Object.keys(TABLE_EDGES).map(edge => [edge, DEFAULT_BORDER_LINE]));
        }

        const own = this.readBorders(tblPr && childElement(tblPr, 'tblBorders'), TABLE_EDGES);
        const edges = Object.assign({}, base, own);
        for (const edge of Object.keys(TABLE_EDGES)) edges[edge] = edges[edge] || null;
        edges.cells = cells;
        return edges;
    }

    /**
     * Give each cell of a parsed table the line each of its four edges draws, as cell.borders
     * (null for no line).
     * A cell's outer edges come from the table's top, bottom, left and right edges and the
     * edges between cells from insideH and insideV; an edge the table style's cell borders
     * set overrides that, and an edge the cell's own w:tcBorders sets overrides both. An edge two
     * cells share then has a line when either cell's edge does, and the heavier of their lines
     * is set on both (see heavierBorderLine). Cells above and below each other share an edge
     * where their grid columns (w:gridSpan included) overlap; such a line is set on a cell only
     * when it runs along that cell's whole edge, so it doesn't reach past a narrower cell along
     * a wider one.
     * @param {Object} table - Parsed table, its rows' cells without borders yet
     * @param {Object<string,?Object>} edges - The table's edges (see tableBorders)
     * @param {Array<Array<Element>>} cellElements - The w:tc element of each cell, row by row
     */
    resolveCellBorders(table, edges, cellElements) {
        const { rows } = table;
        const columns = rows.map(row => {
            let column = 0;
            return row.cells.map(cell => {
                const start = column;
                column += Math.max(1, cell.colspan || 1);
                return [start, column];
            });
        });

        rows.forEach((row, i) => row.cells.forEach((cell, j) => {
            const tcPr = childElement(cellElements[i][j], 'tcPr');
            cell.borders = Object.assign({
                top: i === 0 ? edges.top : edges.insideH,
                right: j === row.cells.length - 1 ? edges.right : edges.insideV,
                bottom: i === rows.length - 1 ? edges.bottom : edges.insideH,
                left: j === 0 ? edges.left : edges.insideV
            }, edges.cells, this.readBorders(tcPr && childElement(tcPr, 'tcBorders'), CELL_EDGES));
        }));

        // Each shared edge is decided from the cells' own edges, so a line doesn't spread
        // along a row through a cell that spans several columns
        const own = rows.map(row => row.cells.map(cell => ({ ...cell.borders })));
        rows.forEach((row, i) => row.cells.forEach((cell, j) => {
            const right = row.cells[j + 1];
            const shared = right && heavierBorderLine(own[i][j].right, own[i][j + 1].left);
            if (shared) {
                cell.borders.right = right.borders.left = shared;
            }
            if (i + 1 === rows.length) return;
            const [start, end] = columns[i][j];
            rows[i + 1].cells.forEach((below, k) => {
                const [belowStart, belowEnd] = columns[i + 1][k];
                const shared = belowStart < end && start < belowEnd &&
                    heavierBorderLine(own[i][j].bottom, own[i + 1][k].top);
                if (shared) {
                    if (belowStart <= start && end <= belowEnd) cell.borders.bottom = shared;
                    if (start <= belowStart && belowEnd <= end) below.borders.top = shared;
                }
            });
        }));
    }

    /**
     * Parse a table cell's properties. Its content is added by parseTable.
     * @param {Element} tc - The table cell element
     * @returns {Object} Parsed cell, without content
     */
    parseTableCell(tc) {
        const cell = {
            type: 'cell',
            blocks: [],
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

        return cell;
    }

    /**
     * Parse the content of a cell of the table being parsed (at this.tableDepth): its paragraphs
     * and the tables nested in it, in document order. A nested table past MAX_TABLE_DEPTH is
     * flattened into the cell's paragraphs in reading order, and reported.
     * @param {Element} tc - The table cell element
     * @returns {Array<Object>} Parsed paragraphs and tables
     */
    parseCellBlocks(tc) {
        const blocks = [];
        for (const block of this.parseBlocks(tc, true)) {
            if (block.type !== 'table' || this.tableDepth < MAX_TABLE_DEPTH) {
                blocks.push(block);
                continue;
            }
            const paragraphs = this.flattenToParagraphs([block]);
            const firstText = paragraphs.map(paragraphText).find(text => text.trim() !== '');
            this.addFinding('nestedTables', shortExample(firstText || ''));
            blocks.push(...paragraphs);
        }
        return blocks;
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
                        // Cell paragraphs are already flattened, nested tables included
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
