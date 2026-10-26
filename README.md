# Word to UDF Converter

Convert Word documents (.docx) to UYAP UDF format - a proprietary document format used by Turkey's National Judiciary Informatics System.

## Features

- 🔒 **100% Client-Side**: Files never leave your browser
- 📄 **DOCX Support**: Converts standard Word documents
- 📋 **Formatting**: Preserves bold, italic, fonts, and alignment
- 📊 **Tables**: Handles table structure with cells and borders
- 🖼️ **Images**: Embeds images as base64
- 🇹🇷 **Turkish UI**: The page is in Turkish by default, with a TR | EN toggle for English (or open it with `?lang=en`)

## Usage

1. Open `index.html` in a web browser
2. Drag and drop a `.docx` file (or click to browse)
3. Click "UDF’YE DÖNÜŞTÜR" ("Convert to UDF" in English)
4. Download the converted `.udf` file

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

### Conversion Pipeline
1. **Parse DOCX**: Extract and parse `word/document.xml`
2. **Extract Elements**: Paragraphs, tables, images
3. **Build Content**: Concatenate all text with offset tracking
4. **Generate XML**: Create UDF-compatible XML structure
5. **Package**: Create ZIP file with `.udf` extension

## Limitations

- Nested tables are flattened: their text goes into the containing cell, in reading order
- Embedded OLE objects are not supported
- Track changes / comments are not preserved (inserted text is kept, deleted text is dropped)
- Headers/footers not yet implemented

## License

MIT — see [LICENSE](LICENSE).

## Resources

- [UDF Format Documentation](https://github.com/saidsurucu/UDF-Toolkit/blob/main/Docs.md)
- [UYAP Document Editor](https://www.uyap.gov.tr/) - Official editor for .udf files
