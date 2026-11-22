# Word to UDF Converter

Convert Word documents (.docx) to UYAP UDF format - a proprietary document format used by Turkey's National Judiciary Informatics System.

## Features

- 🔒 **100% Client-Side**: Files never leave your browser
- 📄 **DOCX Support**: Converts standard Word documents
- 📋 **Formatting**: Preserves bold, italic, fonts, and alignment
- 📊 **Tables**: Handles table structure with cells and borders
- 🖼️ **Images**: Embeds images as base64, including legacy VML pictures (`w:pict`, sized from the shape's style and any `v:group` holding it) and the preview pictures of embedded OLE objects (`w:object`), such as an Excel table or an Equation Editor formula. Charts, SmartArt and shapes without a text box aren't converted; they are listed in the conversion notes
- 📐 **Page Setup**: Keeps the page margins, orientation and header/footer offsets
- 🗂️ **Headers and footers**: Keeps the text of the last section's main header and footer, such as an office name or a file reference (page numbers aren't converted yet)
- ↕️ **Spacing**: Keeps the space before and after paragraphs and their line spacing
- ☑ **Symbols**: Writes check boxes, check and cross marks, bullets, arrows and Greek letters inserted from the Symbol, Wingdings and Wingdings 2 fonts as their Unicode characters, and lists them in a conversion note so they can be checked in the UYAP editor
- ➗ **Equations**: Writes Word equations as plain linear text where they stand, close to Word's own linear format (`A/365`, `x^2`, `√(x)`, `∑_(i=1)^n i`), with the paragraph's formatting, and lists them in a conversion note so they can be checked in the UYAP editor
- ⇥ **Tab stops**: Keeps a paragraph's custom tab stops (left, center, right and decimal, with dot or dash leaders), so lines like `DAVACI : Ad Soyad` stay aligned
- 🔢 **Lists**: Keeps numbered and bulleted lists, including lists that come from a paragraph style (such as numbered headings)
- 📝 **Footnotes and endnotes**: Keeps the reference marks as superscript numbers, and the note text with its run formatting (bold, italic, underline, strikethrough, color, highlighting, font and size, resolved through the document defaults and styles such as `FootnoteText`, as in the body), paragraphs, tabs and line breaks. The space or tab Word puts after the note mark is dropped, and text without a size anywhere is 10 pt. Footnotes are placed before the next manual page break, or at the end; endnotes are moved to the end of the document, after a separator line, numbered the way the document numbers them (`i`, `ii`, ... by default)
- 🇹🇷 **Turkish UI**: The page is in Turkish by default, with a TR | EN toggle for English (or open it with `?lang=en`)

## Usage

1. Open `index.html` in a web browser
2. Drag and drop a `.docx` file (or click to browse)
3. Click "UDF’YE DÖNÜŞTÜR" ("Convert to UDF" in English)
4. Download the converted `.udf` file

While a conversion runs, the page says which step it is on (reading the Word document, then building the UDF) and announces it to screen readers; when it finishes, focus moves to the download button. Removing the file (KALDIR / REMOVE) during a conversion cancels it: its result is discarded and the page goes back to the empty upload state.

After a successful conversion, the page lists any [conversion notes](#conversion-notes): parts of the document the converter dropped or simplified, with what to change in Word. Above the download button it names the source the result came from: the file's name, size and last-modified time, when the conversion finished, and the `.udf` name the download will use (times in the page language). After fixing something in Word, press KAYNAĞI DEĞİŞTİR (REPLACE SOURCE) and choose the new revision, even under the same name: the old download, source line and notes go at once, and the new file converts straight away. Closing the file dialog without a choice keeps the current result, and so does choosing a file that isn't a `.docx` (the page says why). "Başka bir dosya dönüştür" (Convert another file) still clears everything and goes back to the empty upload state.

**Keyboard and screen readers.** The whole loop works without a mouse, and focus never drops back to the top of the page. After a file is chosen (with the file input or by dropping it), focus moves to the convert button. While the conversion runs it is on KALDIR / REMOVE, which cancels it; removing the file, before or during a conversion, puts focus back on the file input. On success focus moves to the download button, and a status line says once, in the page language, that the conversion is complete and how many conversion notes follow ("Dönüştürme tamamlandı. Bu dönüşümle ilgili 3 not var." / "Conversion complete. 3 notes on this conversion."; with no notes, just the first sentence). The notes are a labelled region to read at your own pace, not announced on their own. On failure the error is announced and focus returns to the convert button. "Başka bir dosya dönüştür" (Convert another file) puts focus back on the file input. Focus only moves when the control that had it disappears: if you move it elsewhere during a conversion, for example to the language toggle, it stays there. The upload zone shows the page's dotted focus outline while its file input has keyboard focus.

Password-protected `.docx` files and Word 97–2003 `.doc` files (even renamed to `.docx`) are refused with a message saying how to fix them, and so is a `.docx` whose main document part is damaged: none of them is turned into an empty `.udf`. A `.docx` whose body has content embedded by another program (`w:altChunk`) and nothing else, apart from empty paragraphs, is refused too (`embeddedOnly`), with a message to open it in Word, save it again and convert the saved file: Word turns embedded content into ordinary text.

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

The generator's vocabulary is checked against the format documentation ([`Docs.md` at commit `15496404`](https://github.com/saidsurucu/UDF-Toolkit/blob/15496404cb38ee377ef087cb9b57649c352ed074/Docs.md)): `tests/vocabulary.test.js` converts a corpus of documents and fails on any element or attribute name the documentation doesn't give. Two `<content>` attributes are written beyond it: `superscript="true"` on footnote and endnote marks and note numbers, and `strikethrough="true"` for struck-through text (`w:strike`). Like the documented `<content>` formatting attributes (`family`, `size`, `bold`, …), they are Java Swing `StyleConstants` key names, but they are unverified until checked in the official editor.

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

### Run text
A run's text is written as Word shows it:
- Hidden text (`w:vanish`) is left out, together with any tabs, breaks, pictures and note references in the hidden run, and reported as a `hiddenText` conversion note. Like bold and italic, it resolves through `w:docDefaults`, the paragraph style chain, the character style chain and the run itself, so `<w:vanish w:val="0"/>` on a run shows text a style hides. The paragraph itself is still written, even when all of its text is hidden. Hidden runs in footnotes and endnotes are left out the same way.
- All Caps (`w:caps`) and Small Caps (`w:smallCaps`) text is written in capitals, resolved the same way. Text whose language (`w:lang w:val`) is Turkish, or not set, is upper-cased the Turkish way (`i` becomes `İ`, `ı` becomes `I`); text in another language uses the general rules. Small capitals are written full size: the UDF has no small-caps attribute.
- Highlighting (`w:highlight`) and run shading (`w:shd` in a run's `w:rPr`) are written as the `<content>` element's `background` color, a signed ARGB integer like `foreground`: yellow highlighting is `background="-256"`. Each resolves through `w:docDefaults`, the paragraph style chain, the character style chain and the run itself, and highlighting takes precedence over shading, as Word draws it on top. Highlight names map to Word's fixed palette (`darkYellow` is `808000`, `lightGray` is `C0C0C0`, …); `<w:highlight w:val="none"/>` turns an inherited highlight off. Shading uses its fill (`w:fill`), or its pattern color (`w:color`) for a `solid` pattern, so a patterned shading such as `pct25` is approximated by its fill; `w:val="nil"` or an `auto` fill turns it off. Runs with neither get no `background`. Paragraph shading (`w:shd` in `w:pPr`) is not converted: the documented `<paragraph>` has no background color.
- A non-breaking hyphen (`w:noBreakHyphen`) becomes a hyphen-minus (`-`). An optional hyphen (`w:softHyphen`) is left out, as Word only shows it at a line break.
- A positional tab (`w:ptab`) becomes a plain tab.
- A symbol Word inserted from a symbol font (`<w:sym w:font="Wingdings" w:char="F0FE"/>`, as Insert › Symbol and check boxes write it) becomes its Unicode character, with the run's formatting, wherever it is: body, table cells, headers and footers, footnotes and endnotes. `w:char` is read as hex, in its private-use form (`F0FE`) or plain form (`00FE`), and the font name is matched ignoring case; a `w:sym` without `w:font` uses the run's font. The character is not upper-cased by All Caps, since Word draws the symbol's own glyph. The tables (`SYMBOL_FONTS` in `js/docx-parser.js`) cover:
  - **Symbol**: every code in Unicode's [`symbol.txt`](https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt) mapping outside its Corporate Use Subarea, such as the Greek letters, math signs like `≤`, `≥`, `±`, `×` and `≠`, arrows and the bullet `•`.
  - **Wingdings** and **Wingdings 2**: the empty, checked and crossed boxes (`□`, `☑`, `☒`), check and cross marks (`✓`, `✗`), bullets and arrows, each taken from the Unicode Wingdings proposal [L2/11-052R](https://www.unicode.org/L2/L2011/11052r-wingding.pdf). Wingdings `F0FE` is `☑` and `F0A8` is `□`.
  - Any other font (such as Webdings), or a code its table doesn't list, has no equivalent: nothing is written for it.
  - Every symbol, written or not, is reported as a `symbols` conversion note. A symbol in a hidden run is left out and reported as `hiddenText` instead.
- An equation (Office Math `m:oMath`, also each one in an `m:oMathPara`) becomes plain text where it stands, in a linear form close to Word's own, wherever it is: body, table cells, text boxes, headers and footers, footnotes and endnotes. The UDF has no construct for 2D math, so there is no equation layout or math font.
  - A fraction (`m:f`) is `num/den`; scripts are `base^sup`, `base_sub` and `base_sub^sup`; a radical (`m:rad`) is `√(x)`, or `3√(x)` with a degree that isn't hidden; a delimiter (`m:d`) is its parts between its begin and end characters (`(` and `)` by default), joined by its separator (`|` by default), e.g. `[a|b]`; an n-ary operator (`m:nary`) is the operator (`∫` by default), its limits as `_sub` and `^sup` unless hidden, a space and its body, e.g. `∑_(i=1)^n i`; a function (`m:func`) is its name, a space and its argument, e.g. `sin x`. Any other structure (accents, bars, boxes, matrices, equation arrays, limits, …) is its text in document order.
  - A fraction part, base, script, degree or limit is wrapped in parentheses unless it is a single run of letters or digits, a decimal number, or already one delimiter: `A/365`, `(a+b)/2`, `x^(n+1)`.
  - The text takes the paragraph's run formatting, as a run without its own `w:rPr` would; math fonts such as Cambria Math are not carried over, and All Caps doesn't change it. Each equation after the first in an `m:oMathPara` starts a new line. Deleted runs inside an equation are left out.
  - Every equation is reported as an `equations` conversion note. A hidden one is left out and reported as `hiddenText` instead.

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
- **Table borders.** Each of a table's six edges (`top`, `left`/`start`, `bottom`, `right`/`end`, `insideH`, `insideV`) is resolved on its own: the table style (`w:tblStyle`, else the document's default table style, normally "Normal Table", which has no borders) and its `w:basedOn` chain, then the table's own `w:tblBorders`; the nearest layer that sets an edge wins, and an edge whose `w:val` is `none` or `nil`, or that nothing sets, has no line. A cell's outer edges come from the table's outer edges and the edges between cells from `insideH`/`insideV`; the table style's cell borders (`w:tcPr/w:tcBorders` in the style, through the same chain) override them, and a cell's own `w:tcBorders` overrides both. An edge two cells share is drawn when either cell draws it, and is then set on both with the heavier of the two lines (see **Border lines**), except that a line along only part of a wider cell's edge (`w:gridSpan`) is set only on the narrower cell, so it doesn't run across the whole wide cell. In the UDF, a table with no visible edge is `border="borderNone"`; otherwise it is `borderCell`, a cell with all four edges gets no `borderSpec`, a cell with some gets `borderSpec` with their bits (1 top, 2 right, 4 bottom, 8 left), and a cell with none gets `border="borderNone"`. A document without a styles part keeps a full grid of Word's default line unless the table's own `w:tblBorders` sets edges, since its default table style isn't known.
- **Border lines.** Each edge's line is read from its `w:val`, `w:sz` and `w:color`, through the same layers as the edges. The format has one line per cell, so a cell that draws at least one edge gets the heaviest of the lines it draws (after a vertical merge drops the edges between its cells), as `borderStyle`, `borderWidth` and `borderColor`. The heavier of two lines is the wider one, and on equal width `double` > solid > dashed > dotted; the same rule decides a shared edge. This is a simplified form of Word's conflict rule, which ranks every line style and also compares colors.
  - Style: `single` and `thick` become `borderStyle-solid`, `dotted` `borderStyle-dotted`, `dashed`, `dashSmallGap`, `dotDash`, `dotDotDash` and `dashDotStroked` `borderStyle-dashed`, and `double` `borderStyle-double`. Any other style (`triple`, `thinThick*`, `thickThin*`, `wave`, `doubleWave`, `threeDEmboss`, `threeDEngrave`, `outset`, `inset`, art borders) is approximated as `borderStyle-solid`.
  - Width: `w:sz` is in eighths of a point and is written in points with at least one decimal place (`4` → `0.5`, `8` → `1.0`, `18` → `2.25`). A missing `w:sz`, or one that isn't a positive whole number, counts as `4`.
  - Color: `w:color` (`RRGGBB`) is written as a signed ARGB integer with full alpha (`FF0000` → `-65536`); `auto`, a missing value or one that isn't a hex color is black (`-16777216`). Theme colors (`w:themeColor`, `w:themeShade`, `w:themeTint`) are not resolved: `w:color`, which Word writes alongside them, is used as written.
  - The three attributes are left out when the cell's line is Word's default line (`single`, ½ pt, automatic or black), so an ordinary grid table is written exactly as before and the editor draws its own default line.
  - A cell that draws an approximated style, or whose drawn edges differ in written style, width or color, gets one `tableBorders` conversion note.
- Not read: conditional table-style formatting (`w:tblStylePr`, e.g. first-row borders), cell spacing, per-edge line styles (one line per cell, see above), and theme colors. How the UYAP editor draws `borderSpec` on cells outside a vertical merge has not been checked in the official editor yet.

### Conversion notes
`DocxParser.parse()` also returns `findings`: one `{ code, count, examples, occurrences }` entry per kind of content the conversion drops or simplifies, in a fixed order, and `[]` when nothing applies. `examples` holds up to 3 short location hints (at most about 40 characters each). `occurrences` holds up to 5 `{ example, location }` entries, in the order they were found, each pair of example and location once (see [Locations](#locations)). The page shows them under the download button as "Bu dönüşümle ilgili notlar" / "Notes on this conversion", each with a sentence in Turkish or English saying what happened and what to change in Word. These constructs are reported:

| Code | Reported when | Examples |
| --- | --- | --- |
| `headersFooters` | Header or footer content is not converted (see [Headers and footers](#headers-and-footers)): a section references a first-page or even-page header or footer with non-whitespace text, or a header or footer with text other than the final section's default one; or the converted default header or footer holds a field, picture, text box or table, which is left out. Reported once for the document. | none |
| `comments` | The comments part has comments with text; `count` is the number of such comments. | none |
| `hiddenText` | Text formatted as hidden (`w:vanish`, see [Run text](#run-text)) is left out. Counted per paragraph, header and footer paragraphs and table cells included, and per paragraph of a footnote or endnote once per converted reference. | The start of the paragraph's hidden text |
| `nestedTables` | A table inside a table cell is flattened into that cell. | The start of the nested table's first non-empty paragraph |
| `missingImage` | A picture's relationship or media part is missing, so it is skipped. | The text of the nearest paragraph |
| `unsupportedImage` | A picture's data is not PNG or JPEG, judged by its leading bytes, not its file extension. It is still embedded as-is. | The format: `EMF`, `WMF`, `GIF`, `TIFF`, `SVG` or `unknown` |
| `unconvertedDrawing` | A drawing that shows no picture and holds no text box is left out: a DrawingML chart (`c:chart`), SmartArt diagram or shape (`w:drawing` with no `a:blip`), or a VML shape or OLE object with no preview picture (`w:pict` or `w:object` with no `v:imagedata`). A text box's paragraphs are kept, so a shape that holds one isn't reported. Counted per drawing. | The text of the nearest paragraph |
| `notesMoved` | A footnote reference is converted: its note is written before the next manual page break or at the end, not at the bottom of its page. Counted per reference. | The start of the note text |
| `endnotesMoved` | An endnote reference is converted in a document whose endnotes Word would place or number differently: at the end of each section (`w:endnotePr/w:pos w:val="sectEnd"`) or restarted at each section (`w:numRestart w:val="eachSect"`) when there is more than one section, or restarted at each page. The restart is read from each section's `w:sectPr`, then from `word/settings.xml`; the placement only from `word/settings.xml`, because a `w:pos` in a section's `w:endnotePr` is ignored (ECMA-376 Part 1, 17.11.22). Endnotes at the end of the document with continuous numbering, Word's default, are not reported. Counted per reference. | The start of the note text |
| `lineSpacing` | A converted paragraph's line rule, resolved from `w:docDefaults`, the paragraph style chain and the paragraph itself, is `exact` or `atLeast`, so it gets single spacing. Counted per paragraph, table cells included. | The start of the paragraph's text |
| `tabStops` | A converted paragraph's resolved tab stops include a `bar` stop, or a stop with an `underscore` or `heavy` leader, which becomes no leader. A stop the paragraph clears or replaces doesn't count. Counted per paragraph. | The start of the paragraph's text |
| `symbols` | A symbol (`w:sym`, see [Run text](#run-text)) is converted: it is written as its Unicode character, which the UYAP editor's fonts may draw differently, or, with no equivalent, left out. Counted per symbol, in the body, table cells, headers and footers, and in footnotes and endnotes once per converted reference. | The font and code, then `→` and the character written, e.g. `Wingdings F0FE → ☑`; a symbol left out has no arrow, e.g. `Webdings F021` |
| `tableBorders` | A converted table cell's border lines can't all be written (see [Tables](#tables)): an edge it draws has a line style with no documented equivalent, approximated as solid, or its drawn edges differ in style, width or color, so its heaviest line was written for the whole cell. Counted per cell; cells of nested tables, which are flattened, aren't counted. | The start of the cell's text |
| `equations` | An equation (`m:oMath`, see [Run text](#run-text)) is written as plain linear text. Counted per equation, in the body, table cells, text boxes, headers and footers, and in footnotes and endnotes once per converted reference. | The equation's linear text, e.g. `A/365` |
| `embeddedContent` | Content another program embedded in the file (`w:altChunk`: HTML, RTF, MHT or another `.docx`, as web apps and document-assembly tools export it) is left out: it isn't converted. Counted per chunk, in the body, block-level content controls, table cells and text boxes; a chunk in the converted header or footer is reported as `headersFooters`. A body whose only content is such chunks is refused instead (see [Usage](#usage)). | The embedded part's file name, e.g. `afchunk.mht` |

`lineSpacing` and `tabStops` count the paragraphs that are written as paragraphs. A body paragraph that holds only manual page breaks becomes the page breaks alone, so it is not counted; in a table cell a page break is a line break, and such a paragraph is written and counted. Nothing is reported for the content of a cell that continues a vertical merge, which is written empty.

The notes only report: the converted `.udf` is the same with or without them.

#### Locations
Each occurrence says where it is, so the user can find it in Word. `location` is one of:

- `{ kind: 'paragraph', paragraph: 12, heading: 'AÇIKLAMALAR' }`: the 12th paragraph of the body. Body paragraphs are counted from 1 in document order, empty ones included, so the number matches the ¶ marks Word shows; content controls (`w:sdt`), custom XML and `mc:AlternateContent` are looked through. Paragraphs in table cells and text boxes are not counted: what a text box reports takes the place of the body paragraph that anchors it. An embedded chunk (`w:altChunk`) in the body counts as one paragraph, where its content would start.
- `{ kind: 'cell', table: 2, row: 3, cell: 1, heading: 'DELİLLER' }`: the 2nd table of the body, its 3rd row (`w:tr`) and the 1st cell (`w:tc`) of that row, all counted from 1. What a nested table reports takes the place of the outer table's cell.
- `{ kind: 'header' }` or `{ kind: 'footer' }`: the converted header or footer.
- `{ kind: 'document' }`: the whole document. `headersFooters`, `comments` and `endnotesMoved` have no single place and always use it.

What a footnote's or endnote's text reports (`symbols`, `equations`, `hiddenText`), and `notesMoved`, takes the place of the note's reference, plus `note: { kind: 'footnote' }` or `note: { kind: 'endnote' }`. A picture's `missingImage` or `unconvertedDrawing` takes the place of the paragraph or cell holding it, and keeps the nearest paragraph's text as its example.

`heading` is the text (shortened like an example) of the last body paragraph before the place that is a heading, and is left out when there is none. A body paragraph with text is a heading when its outline level (`w:outlineLvl`, from the paragraph's own `w:pPr`, else its paragraph style chain, as the built-in Heading 1–9 styles set it) is 0 to 8. Without such a paragraph before the place, a short petition heading is used instead: a body paragraph of at most 80 characters whose text runs are all bold, or whose letters are all upper case (`AÇIKLAMALAR :`, `SONUÇ VE İSTEM`). Paragraphs in tables are never headings, and a heading's own findings name the heading before it. The bold and upper-case rule is a hint and may name a line that isn't a heading; the paragraph or cell number is exact.

The page lists a note's places, one line each, under "Nerede:" / "Where:", e.g. `Paragraf 12 · “AÇIKLAMALAR” başlığı altında · “metin örneği”` / `Paragraph 12 · under “AÇIKLAMALAR” · “text excerpt”`, `Tablo 2, satır 3, hücre 1` / `Table 2, row 3, cell 1`, `Üst bilgi` / `Header`, `Dipnot · Paragraf 7` / `Footnote · Paragraph 7`. Document-wide places are not listed; a note with no other place lists its examples as before. A sentence under the notes title explains how paragraphs are counted. "Sayı: N" / "Count: N" still gives the total, which can be more than the places listed.

### Conversion Pipeline
1. **Parse DOCX**: Extract and parse `word/document.xml`
2. **Extract Elements**: Paragraphs, tables, images
3. **Build Content**: Concatenate all text with offset tracking
4. **Generate XML**: Create UDF-compatible XML structure
5. **Package**: Create ZIP file with `.udf` extension

## Limitations

- Nested tables are flattened: their text goes into the containing cell, in reading order
- Embedded OLE objects are written as their preview picture (often EMF or WMF, which UDF doesn't define, so they are reported as `unsupportedImage`); charts, SmartArt diagrams and shapes without a text box are left out and listed in the conversion notes. VML pictures are written inline where their run is, without cropping, wrapping or position
- Documents whose parts decompress to more than 100 MiB each, or 300 MiB together, are refused (see [Usage](#usage))
- Password-protected documents and `.doc` files are not converted: remove the password, or save the file as `.docx`, in Word first
- Content embedded by another program (`w:altChunk`) is not converted: it is listed in the conversion notes, and a document with nothing else is refused. Saving the file again in Word turns it into ordinary text
- Equations are written as plain linear text, without 2D layout or a math font, and listed in the conversion notes
- Track changes / comments are not preserved (inserted text is kept, deleted text is dropped)
- Only the final section's default header and footer are converted, as plain paragraphs: page numbers and other fields, pictures, text boxes and tables in them, first-page and even-page headers and footers, and other sections' headers and footers are left out and listed in the conversion notes
- Footnote and endnote text keeps its run formatting and paragraphs, but not paragraph alignment, indents or spacing inside notes (note paragraphs are left-aligned, without indents), nor tab stops in notes; pictures, page breaks and tables inside notes are not converted as such (a table's text is kept as paragraphs). All endnotes go at the end of the document, even when Word places them at the end of each section, and they are always numbered from the start. Footnotes and such endnotes are listed in the conversion notes
- Exact and at-least line heights, automatic paragraph spacing and contextual spacing are not converted (see Paragraph spacing); exact and at-least line heights are listed in the conversion notes
- Bar tab stops and underscore or heavy tab leaders are not converted (see Tab stops), and are listed in the conversion notes
- Hidden text is left out (see Run text) and listed in the conversion notes; Small Caps text is written in full-size capitals

## License

MIT — see [LICENSE](LICENSE).

## Resources

- [UDF Format Documentation](https://github.com/saidsurucu/UDF-Toolkit/blob/main/Docs.md)
- [UYAP Document Editor](https://www.uyap.gov.tr/) - Official editor for .udf files
