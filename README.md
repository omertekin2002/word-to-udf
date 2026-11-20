# Word to UDF Converter

Convert Word documents (.docx) to UYAP UDF format - a proprietary document format used by Turkey's National Judiciary Informatics System.

## Features

- 🔒 **100% Client-Side**: Files never leave your browser
- 📄 **DOCX Support**: Converts standard Word documents
- 📋 **Formatting**: Preserves bold, italic, fonts, and alignment
- 📊 **Tables**: Handles table structure with cells and borders
- 🖼️ **Images**: Embeds images as base64
- 📐 **Page Setup**: Keeps the page margins, orientation and header/footer offsets
- 🗂️ **Headers and footers**: Keeps the text of the last section's main header and footer, such as an office name or a file reference (page numbers aren't converted yet)
- ↕️ **Spacing**: Keeps the space before and after paragraphs and their line spacing
- ⇥ **Tab stops**: Keeps a paragraph's custom tab stops (left, center, right and decimal, with dot or dash leaders), so lines like `DAVACI : Ad Soyad` stay aligned
- 🔢 **Lists**: Keeps numbered and bulleted lists, including lists that come from a paragraph style (such as numbered headings)
- 📝 **Footnotes and endnotes**: Keeps the reference marks as superscript numbers and the note text. Footnotes are placed before the next manual page break, or at the end; endnotes are moved to the end of the document, after a separator line, numbered the way the document numbers them (`i`, `ii`, ... by default)
- 🇹🇷 **Turkish UI**: The page is in Turkish by default, with a TR | EN toggle for English (or open it with `?lang=en`)

## Usage

1. Open `index.html` in a web browser
2. Drag and drop a `.docx` file (or click to browse)
3. Click "UDF’YE DÖNÜŞTÜR" ("Convert to UDF" in English)
4. Download the converted `.udf` file

While a conversion runs, the page says which step it is on (reading the Word document, then building the UDF) and announces it to screen readers; when it finishes, focus moves to the download button. Removing the file (KALDIR / REMOVE) during a conversion cancels it: its result is discarded and the page goes back to the empty upload state.

After a successful conversion, the page lists any [conversion notes](#conversion-notes): parts of the document the converter dropped or simplified, with what to change in Word.

Password-protected `.docx` files and Word 97–2003 `.doc` files (even renamed to `.docx`) are refused with a message saying how to fix them, and so is a `.docx` whose main document part is damaged: none of them is turned into an empty `.udf`.

A `.docx` is also refused, with a message to save it again in Word or compress or remove large pictures, when one of its parts would decompress to more than 100 MiB, or all the parts a conversion reads to more than 300 MiB together (`MAX_PART_BYTES` and `MAX_TOTAL_BYTES` in `js/docx-parser.js`). Each part's declared size is checked before it is decompressed, and decompression stops as soon as a part produces more than it declares or the limits allow, so a zip bomb or a damaged size field gets an error message instead of running the tab out of memory. A petition is far below both limits. Only the pictures the document shows are read: an image part it relates to but doesn't use (left over from a deleted picture, say) is never decompressed.

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

`bun run test` runs `tests/*.test.js` and needs no browser. A separate smoke test, `tests/browser/smoke.test.js`, serves the site on `127.0.0.1` and drives a whole conversion through the real page in headless Chromium with [Playwright](https://playwright.dev): the page loading with no CSP violations or failed requests, choosing a file, converting, downloading the `.udf`, the conversion notes, a damaged file, the TR | EN toggle and "convert another". Download Chromium once, then run it:

```bash
bunx playwright install chromium
bun run test:browser
```

The site itself still has no build step; `package.json` exists only for the tests.

`tests/petitions.test.js` converts whole synthetic petitions built in `tests/fixtures/petitions.js`. Known conversion gaps show up as `todo` tests: they are listed in the output but don't fail the run. The fixtures use invented placeholder content only (`Ad Soyad`, `00000000000`, `2026/0000`); never add real names, ID numbers, case numbers or addresses to them.

## Privacy and dependencies

The page loads nothing from another site. [JSZip](https://stuk.github.io/jszip/) is vendored as `js/vendor/jszip.min.js`, a copy of `dist/jszip.min.js` from the `jszip` devDependency, which is pinned to an exact version. To update it, change that version, run `bun install`, and copy `node_modules/jszip/dist/jszip.min.js` over the vendored file; `tests/security.test.js` fails if the two differ.

A Content-Security-Policy `<meta>` tag in `index.html` restricts scripts, styles and images to the same origin, blocking cross-origin resource loads, script connection APIs such as `fetch`, XMLHttpRequest and WebSocket (`connect-src 'none'`), and form submissions. Same-origin image requests remain permitted, so the policy alone does not guarantee that document contents stay local; the converter's client-side code processes documents without uploading them. Inline scripts, inline `style` attributes and `on*=` handlers would be blocked by it; the same test file checks that the page has none.

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

### Headers and footers
The default header and footer (`w:headerReference` / `w:footerReference` with `w:type="default"`) of the document's final section, the one the page setup comes from, are converted. A section with no default header or footer of its own uses the previous section's, as in Word.
- Their top-level paragraphs, including those in content controls, are read like body paragraphs: styles, fonts, emphasis, alignment, indents, spacing, tab stops, tabs and line breaks resolve the same way.
- They are written as `<header>`, the first child of `<elements>`, and `<footer>`, the last one, each holding `<paragraph>` elements. Their text goes in the shared content: header text first, then the body, footnotes and endnotes, then footer text, as the UDF documentation describes.
- No `pageNumber-*` attributes are written: the `pageNumber-spec` format isn't documented well enough to produce.
- Fields (`w:fldSimple`, `w:fldChar`, such as `PAGE` and `NUMPAGES`) are left out together with their displayed result, which would otherwise be frozen as static text. Pictures, text boxes and tables in a header or footer are left out too, and so are page breaks.
- A header or footer with no non-whitespace text left writes no element.
- Not converted, and reported as a `headersFooters` conversion note: first-page (`w:type="first"`) and even-page (`w:type="even"`) headers and footers with text; an earlier section's header or footer with text that differs from the final section's; and fields, pictures, text boxes and tables in the converted header or footer. The UDF format documents a single header and footer.

### Paragraph spacing
Paragraph spacing (`w:spacing`) is read from the document defaults, the paragraph style and its `w:basedOn` chain, and the paragraph itself, each attribute on its own: a style can set the space before while the paragraph sets only the space after.
- `w:before` and `w:after` become `SpaceAbove` and `SpaceBelow`, in points.
- `w:line` with `w:lineRule="auto"` (or no rule) becomes `LineSpacing`, the extra fraction over single spacing, to two decimals: `276` (1.15 lines) is `0.15`, `360` (1.5 lines) is `0.5`, `480` (double) is `1.0`.
- Attributes that would be zero, or single (or tighter) line spacing, are left out.
- Exact and at-least line heights (`w:lineRule="exact"` / `"atLeast"`) have no UDF equivalent, so those paragraphs get single spacing, and the page reports them as a `lineSpacing` conversion note.
- Automatic spacing (`w:beforeAutospacing`, `w:afterAutospacing`), `w:contextualSpacing` and spacing in lines (`w:beforeLines`, `w:afterLines`) are ignored.

### Tab stops
A paragraph's tab stops come from `w:tabs` in `w:docDefaults`, the paragraph style and its `w:basedOn` chain, and the paragraph itself. Each layer adds to the stops it inherits: a stop replaces any inherited stop at the same position, and `w:val="clear"` removes the inherited stop at its position. The result is written as the paragraph's `TabSet` attribute, `position:alignment:leader` per stop, comma-separated and in position order, with positions in points to one decimal: `w:pos="1985"` and `w:pos="2268"` become `TabSet="99.3:0:0,113.4:0:0"`.
- Alignment: `left`, `start` and `num` are `0`, `center` is `1`, `right` and `end` are `2`, `decimal` is `3`.
- Leader: `dot` and `middleDot` are `1`, `hyphen` is `2`; `none`, `underscore`, `heavy` or no leader are `0`.
- `bar` stops (a vertical line, not a text position) and stops with an unreadable or negative position are dropped. A paragraph left with no stops gets no `TabSet`. A paragraph that ends up with a `bar` stop, or with an `underscore` or `heavy` leader, gets a `tabStops` conversion note.
- Two stops are at the same position when their `w:pos` values are the same length, in whatever unit each is written (`w:pos="432"` and `w:pos="0.3in"`). Stops a twip apart are different stops: a clear or a stop next to an inherited stop leaves it in place.
- No position is written twice. When two stops that are left round to the same tenth of a point (1441 and 1442 twips are both `72.1`), only the first in position order is written, the one a tab reaches first.
- Paragraphs in table cells get their stops the same way.
- Positions are written as Word measures them, from the left margin. Whether the UYAP editor measures `TabSet` positions from the margin or from the paragraph's left indent is not documented and has not been checked in the official editor, so stops in indented paragraphs may land elsewhere there.
- The document's default tab interval (`w:defaultTabStop`), the implicit stop Word adds at a hanging indent, and tab stops in footnote and endnote text are not converted.

### Lists
A paragraph's list comes from `w:numPr`, read from the paragraph style and its `w:basedOn` chain and from the paragraph itself, each attribute on its own: a numbered paragraph style gives the list (`w:numId`) and level (`w:ilvl`), and the paragraph can set just its own `w:ilvl` to move to another level of that list, or just its own `w:numId`.
- `w:numId w:val="0"`, on the paragraph or on a style nearer in the chain, switches numbering off, including numbering a style would give.
- A missing `w:ilvl` means level 0.
- Bullet or number, and the number style, come from that list level in `word/numbering.xml`. A list missing from `numbering.xml` is still written as a list.
- Numbering in `w:docDefaults` and table styles, restarts (`w:startOverride`) and the level a `w:lvl/w:pStyle` names are not read.

### Tables
- Cell shading (`w:shd`) is written as `fillColor`, and vertical alignment (`w:vAlign`) as `align="vcenter"` or `align="bottom"`; top, the default, is left out.
- The format has no `rowspan`, so a vertical merge (`w:vMerge`) keeps one cell per row: the first holds the text, the rows below get an empty cell in the same grid column, shaded like the first. The cells are joined by dropping the edges between them from their `borderSpec` (in a full grid: `11` first, `10` middle, `14` last).
- **Table borders.** Each of a table's six edges (`top`, `left`/`start`, `bottom`, `right`/`end`, `insideH`, `insideV`) is resolved on its own: the table style (`w:tblStyle`, else the document's default table style, normally "Normal Table", which has no borders) and its `w:basedOn` chain, then the table's own `w:tblBorders`; the nearest layer that sets an edge wins, and an edge whose `w:val` is `none` or `nil`, or that nothing sets, has no line. A cell's outer edges come from the table's outer edges and the edges between cells from `insideH`/`insideV`; the table style's cell borders (`w:tcPr/w:tcBorders` in the style, through the same chain) override them, and a cell's own `w:tcBorders` overrides both. An edge two cells share is drawn when either cell draws it, and is then set on both. In the UDF, a table with no visible edge is `border="borderNone"`; otherwise it is `borderCell`, a cell with all four edges gets no `borderSpec`, a cell with some gets `borderSpec` with their bits (1 top, 2 right, 4 bottom, 8 left), and a cell with none gets `border="borderNone"`. A document without a styles part keeps a full grid unless the table's own `w:tblBorders` removes edges, since its default table style isn't known.
- Not read: conditional table-style formatting (`w:tblStylePr`, e.g. first-row borders), cell spacing, conflict rules by border weight, and border style, width and color. How the UYAP editor draws `borderSpec` on cells outside a vertical merge has not been checked in the official editor yet.

### Conversion notes
`DocxParser.parse()` also returns `findings`: one `{ code, count, examples }` entry per kind of content the conversion drops or simplifies, in a fixed order, and `[]` when nothing applies. `examples` holds up to 3 short location hints (at most about 40 characters each). The page shows them under the download button as "Bu dönüşümle ilgili notlar" / "Notes on this conversion", each with a sentence in Turkish or English saying what happened and what to change in Word. These constructs are reported:

| Code | Reported when | Examples |
| --- | --- | --- |
| `headersFooters` | Header or footer content is not converted (see [Headers and footers](#headers-and-footers)): a section references a first-page or even-page header or footer with non-whitespace text, or a header or footer with text other than the final section's default one; or the converted default header or footer holds a field, picture, text box or table, which is left out. Reported once for the document. | none |
| `comments` | The comments part has comments with text; `count` is the number of such comments. | none |
| `nestedTables` | A table inside a table cell is flattened into that cell. | The start of the nested table's first non-empty paragraph |
| `missingImage` | A picture's relationship or media part is missing, so it is skipped. | The text of the nearest paragraph |
| `unsupportedImage` | A picture's data is not PNG or JPEG, judged by its leading bytes, not its file extension. It is still embedded as-is. | The format: `EMF`, `WMF`, `GIF`, `TIFF`, `SVG` or `unknown` |
| `notesMoved` | A footnote reference is converted: its note is written as plain text before the next manual page break or at the end, not at the bottom of its page. Counted per reference. | The start of the note text |
| `endnotesMoved` | An endnote reference is converted in a document whose endnotes Word would place or number differently: at the end of each section (`w:endnotePr/w:pos w:val="sectEnd"`) or restarted at each section (`w:numRestart w:val="eachSect"`) when there is more than one section, or restarted at each page. The restart is read from each section's `w:sectPr`, then from `word/settings.xml`; the placement only from `word/settings.xml`, because a `w:pos` in a section's `w:endnotePr` is ignored (ECMA-376 Part 1, 17.11.22). Endnotes at the end of the document with continuous numbering, Word's default, are not reported. Counted per reference. | The start of the note text |
| `lineSpacing` | A converted paragraph's line rule, resolved from `w:docDefaults`, the paragraph style chain and the paragraph itself, is `exact` or `atLeast`, so it gets single spacing. Counted per paragraph, table cells included. | The start of the paragraph's text |
| `tabStops` | A converted paragraph's resolved tab stops include a `bar` stop, or a stop with an `underscore` or `heavy` leader, which becomes no leader. A stop the paragraph clears or replaces doesn't count. Counted per paragraph. | The start of the paragraph's text |

`lineSpacing` and `tabStops` count the paragraphs that are written as paragraphs. A body paragraph that holds only manual page breaks becomes the page breaks alone, so it is not counted; in a table cell a page break is a line break, and such a paragraph is written and counted. Nothing is reported for the content of a cell that continues a vertical merge, which is written empty.

The notes only report: the converted `.udf` is the same with or without them.

### Conversion Pipeline
1. **Parse DOCX**: Extract and parse `word/document.xml`
2. **Extract Elements**: Paragraphs, tables, images
3. **Build Content**: Concatenate all text with offset tracking
4. **Generate XML**: Create UDF-compatible XML structure
5. **Package**: Create ZIP file with `.udf` extension

## Limitations

- Nested tables are flattened: their text goes into the containing cell, in reading order
- Embedded OLE objects are not supported
- Documents whose parts decompress to more than 100 MiB each, or 300 MiB together, are refused (see [Usage](#usage))
- Password-protected documents and `.doc` files are not converted: remove the password, or save the file as `.docx`, in Word first
- Track changes / comments are not preserved (inserted text is kept, deleted text is dropped)
- Only the final section's default header and footer are converted, as plain paragraphs: page numbers and other fields, pictures, text boxes and tables in them, first-page and even-page headers and footers, and other sections' headers and footers are left out and listed in the conversion notes
- Footnote and endnote text is converted as plain text, without its formatting; a tab or line break inside a note becomes a space. All endnotes go at the end of the document, even when Word places them at the end of each section, and they are always numbered from the start. Footnotes and such endnotes are listed in the conversion notes
- Exact and at-least line heights, automatic paragraph spacing and contextual spacing are not converted (see Paragraph spacing); exact and at-least line heights are listed in the conversion notes
- Bar tab stops and underscore or heavy tab leaders are not converted (see Tab stops), and are listed in the conversion notes

## License

MIT — see [LICENSE](LICENSE).

## Resources

- [UDF Format Documentation](https://github.com/saidsurucu/UDF-Toolkit/blob/main/Docs.md)
- [UYAP Document Editor](https://www.uyap.gov.tr/) - Official editor for .udf files
