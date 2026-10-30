# Word to UDF Converter

Convert Word documents (.docx) to UYAP UDF format - a proprietary document format used by Turkey's National Judiciary Informatics System.

## Features

- 🔒 **100% Client-Side**: Files never leave your browser
- 📄 **DOCX Support**: Converts standard Word documents
- 📋 **Formatting**: Preserves bold, italic, fonts, and alignment
- 📊 **Tables**: Handles table structure with cells and borders
- 🖼️ **Images**: Embeds images as base64
- 📐 **Page Setup**: Keeps the page margins, orientation and header/footer offsets
- ↕️ **Spacing**: Keeps the space before and after paragraphs and their line spacing
- 🇹🇷 **Turkish UI**: The page is in Turkish by default, with a TR | EN toggle for English (or open it with `?lang=en`)

## Usage

1. Open `index.html` in a web browser
2. Drag and drop a `.docx` file (or click to browse)
3. Click "UDF’YE DÖNÜŞTÜR" ("Convert to UDF" in English)
4. Download the converted `.udf` file

Password-protected `.docx` files and Word 97–2003 `.doc` files (even renamed to `.docx`) are refused with a message saying how to fix them, and so is a `.docx` whose main document part is damaged: none of them is turned into an empty `.udf`.

## Local Development

To run locally, start a simple HTTP server:

```bash
# Using Python 3
python3 -m http.server 8000

# Using Node.js
npx serve .

# Using PHP
php -S localhost:8000
```

Then open `http://localhost:8000` in your browser.

## Running tests

The tests load the browser scripts into Node with [jsdom](https://github.com/jsdom/jsdom) and run whole conversions on small `.docx` files built in memory. They need Node 22+ and [bun](https://bun.sh):

```bash
bun install
bun run test
```

The site itself still has no build step; `package.json` exists only for the tests.

`tests/petitions.test.js` converts whole synthetic petitions built in `tests/fixtures/petitions.js`. Known conversion gaps show up as `todo` tests: they are listed in the output but don't fail the run. The fixtures use invented placeholder content only (`Ad Soyad`, `00000000000`, `2026/0000`); never add real names, ID numbers, case numbers or addresses to them.

## Technical Details

### UDF Format
UDF (UYAP Document Format) is a ZIP archive containing:
- `content.xml` - XML with document structure

The format uses an offset-based content model where all text is stored in a CDATA block, and elements reference text positions using `startOffset` and `length` attributes.

### Page Setup
The `<pageFormat>` element takes its values from the document's final section properties (`w:sectPr`):
- Margins and header/footer offsets come from `w:pgMar`, converted from twips to points. A gutter is added to the left margin (to the top with `w:gutterAtTop`, to the right with `w:rtlGutter`).
- The page is landscape when `w:pgSz` says `w:orient="landscape"` or is wider than it is tall.
- The paper size is always A4, the only size the UDF format documents.
- A value the document doesn't set, or sets to something unreadable, keeps the converter's previous default (left 15 mm, right 10 mm, top and bottom 5 mm, portrait).
- Only the last section's setup is used; per-section page setup isn't supported.

### Paragraph spacing
Paragraph spacing (`w:spacing`) is read from the document defaults, the paragraph style and its `w:basedOn` chain, and the paragraph itself, each attribute on its own: a style can set the space before while the paragraph sets only the space after.
- `w:before` and `w:after` become `SpaceAbove` and `SpaceBelow`, in points.
- `w:line` with `w:lineRule="auto"` (or no rule) becomes `LineSpacing`, the extra fraction over single spacing, to two decimals: `276` (1.15 lines) is `0.15`, `360` (1.5 lines) is `0.5`, `480` (double) is `1.0`.
- Attributes that would be zero, or single (or tighter) line spacing, are left out.
- Exact and at-least line heights (`w:lineRule="exact"` / `"atLeast"`) have no UDF equivalent, so those paragraphs get single spacing.
- Automatic spacing (`w:beforeAutospacing`, `w:afterAutospacing`), `w:contextualSpacing` and spacing in lines (`w:beforeLines`, `w:afterLines`) are ignored.

### Tables
- Cell shading (`w:shd`) is written as `fillColor`, and vertical alignment (`w:vAlign`) as `align="vcenter"` or `align="bottom"`; top, the default, is left out.
- The format has no `rowspan`, so a vertical merge (`w:vMerge`) keeps one cell per row: the first holds the text, the rows below get an empty cell in the same grid column, shaded like the first. In bordered tables the cells are joined by `borderSpec` dropping the edges between them (`11` first, `10` middle, `14` last); cells outside a merge get no `borderSpec`.

### Conversion Pipeline
1. **Parse DOCX**: Extract and parse `word/document.xml`
2. **Extract Elements**: Paragraphs, tables, images
3. **Build Content**: Concatenate all text with offset tracking
4. **Generate XML**: Create UDF-compatible XML structure
5. **Package**: Create ZIP file with `.udf` extension

## Limitations

- Nested tables are flattened: their text goes into the containing cell, in reading order
- Embedded OLE objects are not supported
- Password-protected documents and `.doc` files are not converted: remove the password, or save the file as `.docx`, in Word first
- Track changes / comments are not preserved (inserted text is kept, deleted text is dropped)
- Headers/footers not yet implemented
- Exact and at-least line heights, automatic paragraph spacing and contextual spacing are not converted (see Paragraph spacing)

## License

MIT — see [LICENSE](LICENSE).

## Resources

- [UDF Format Documentation](https://github.com/saidsurucu/UDF-Toolkit/blob/main/Docs.md)
- [UYAP Document Editor](https://www.uyap.gov.tr/) - Official editor for .udf files
