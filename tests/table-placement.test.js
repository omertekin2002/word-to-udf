const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, findingSummaries, loadConverter } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

// Table alignment (w:tblPr/w:jc), indent (w:tblInd) and floating position (w:tblpPr). The documented
// <table> has none of them, so every table is written at the left margin, in the text flow. A table
// Word draws elsewhere is reported as a tablePlacement conversion note; the UDF itself is unchanged.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
// A table: tblPr and rows of cells, each cell's content; one 1500-twip grid column per cell of the first row
const table = ({ tblPr = '', rows }) => '<w:tbl>' +
    (tblPr ? `<w:tblPr>${tblPr}</w:tblPr>` : '') +
    `<w:tblGrid>${rows[0].map(() => '<w:gridCol w:w="1500"/>').join('')}</w:tblGrid>` +
    rows.map(cells => `<w:tr>${cells.map(content => `<w:tc>${content}</w:tc>`).join('')}</w:tr>`).join('') +
    '</w:tbl>';
// The issue's reproduction: a 2-column table between two paragraphs
const twoColumns = (tblPr, first = 'Harç') => table({ tblPr, rows: [[p(r(first)), p(r('500 TL'))]] });
const between = tblPr => p(r('Önce')) + twoColumns(tblPr) + p(r('Sonra'));
const tableStyle = (id, { tblPr = '', basedOn = '', isDefault = false } = {}) =>
    `<w:style w:type="table"${isDefault ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${id}"/>` +
    `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}${tblPr ? `<w:tblPr>${tblPr}</w:tblPr>` : ''}</w:style>`;

const jc = val => `<w:jc w:val="${val}"/>`;
const ind = (w, type = 'dxa') => `<w:tblInd w:w="${w}"${type === null ? '' : ` w:type="${type}"`}/>`;
const FLOATING = '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpXSpec="right" w:tblpY="1"/>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;
const examples = (parsed, code) => occurrences(parsed, code).map(o => o.example);
const count = (parsed, code) => (finding(parsed, code) || { count: 0 }).count;

test('the issue\'s reproduction: centred, right-aligned, indented and floating tables each give one note, and the UDF is unchanged', async () => {
    const plainTable = await convertBody(between(''));
    assert.deepEqual(findingSummaries(plainTable.parsed.findings), []);

    for (const tblPr of [jc('center'), jc('right'), ind(720), FLOATING]) {
        const { xml, parsed } = await convertBody(between(tblPr));
        assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'tablePlacement', count: 1, examples: ['Harç'] }], tblPr);
        assert.deepEqual(occurrences(parsed, 'tablePlacement'),
            [{ example: 'Harç', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }], tblPr);
        assert.match(xml, /<table tableName="Table" columnCount="2" columnSpans="[^"]*" border="borderCell">/);
        assert.equal(xml, plainTable.xml, tblPr);
    }
});

test('end alignment is reported; left, start and no alignment are not', async () => {
    const { parsed } = await convertBody(
        twoColumns(jc('end'), 'Son') + p(r('-')) +
        twoColumns(jc('left'), 'Sol') + p(r('-')) +
        twoColumns(jc('start'), 'Başlangıç') + p(r('-')) +
        twoColumns('', 'Yok'));
    assert.deepEqual(examples(parsed, 'tablePlacement'), ['Son']);
    assert.equal(count(parsed, 'tablePlacement'), 1);
});

test('an indent more than ⅛ in from 0 is reported; Word\'s default indents, nil, pct and auto are not', async () => {
    const indents = [
        ['720', ind(720), true],
        ['eksi 720', ind(-720), true],
        ['türsüz 720', ind(720, null), true],
        ['181', ind(181), true],
        ['180', ind(180), false],
        ['eksi 108', ind(-108), false],
        ['eksi 115', ind(-115), false],
        ['sıfır', ind(0), false],
        ['nil', ind(720, 'nil'), false],
        ['yüzde', ind(720, 'pct'), false],
        ['otomatik', ind(720, 'auto'), false],
        ['bozuk', '<w:tblInd w:w="yarım" w:type="dxa"/>', false]
    ];
    const { parsed } = await convertBody(indents.map(([name, tblPr]) => twoColumns(tblPr, name)).join(p(r('-'))));
    assert.deepEqual(examples(parsed, 'tablePlacement'), indents.filter(i => i[2]).map(i => i[0]));
    assert.equal(count(parsed, 'tablePlacement'), 4);
});

test('alignment and indent resolve through the table style\'s basedOn chain to the table\'s own w:tblPr', async () => {
    const styles = tableStyle('Normal', { isDefault: true, tblPr: ind(0) + '<w:tblCellMar/>' }) +
        tableStyle('Orta', { tblPr: jc('center') }) +
        tableStyle('OrtaAlt', { basedOn: 'Orta' }) +
        tableStyle('Girintili', { tblPr: ind(720) }) +
        tableStyle('GirintisizAlt', { basedOn: 'Girintili', tblPr: ind(-108) });
    const style = id => `<w:tblStyle w:val="${id}"/>`;
    const { parsed } = await convertBody([
        twoColumns(style('Orta'), 'Stil'),
        twoColumns(style('OrtaAlt'), 'Zincir'),
        twoColumns(style('Girintili'), 'Stil girinti'),
        // The table's own left alignment and nil indent override the style's
        twoColumns(style('Orta') + jc('left'), 'Kendi sol'),
        twoColumns(style('Girintili') + ind(0, 'nil'), 'Kendi nil'),
        // A nearer style's default indent overrides the one it is based on
        twoColumns(style('GirintisizAlt'), 'Alt stil'),
        // The default table style, for no style and for an unknown one
        twoColumns('', 'Varsayılan'),
        twoColumns(style('Yok'), 'Bilinmeyen')
    ].join(p(r('-'))), { styles });
    assert.deepEqual(examples(parsed, 'tablePlacement'), ['Stil', 'Zincir', 'Stil girinti']);
    assert.equal(count(parsed, 'tablePlacement'), 3);

    const fromDefault = await convertBody(twoColumns('', 'Varsayılan') + p(r('-')) + twoColumns(jc('left') + ind(0), 'Sol'),
        { styles: tableStyle('Normal', { isDefault: true, tblPr: jc('center') }) });
    assert.deepEqual(examples(fromDefault.parsed, 'tablePlacement'), ['Varsayılan']);
});

test('a table both centred and floating counts once, and a nested centred table once, at the cell that holds it', async () => {
    const nested = table({ tblPr: jc('center'), rows: [[p(r('İç tablo'))]] });
    const { parsed } = await convertBody(
        p(r('DAVALILAR')) +
        twoColumns(jc('center') + FLOATING + ind(720), 'Hem ortalı hem kayan') + p(r('-')) +
        table({ rows: [[p(r('Dış')), p('') + nested]] }));
    assert.equal(count(parsed, 'tablePlacement'), 2);
    assert.deepEqual(occurrences(parsed, 'tablePlacement'), [
        { example: 'Hem ortalı hem kayan', location: { kind: 'cell', table: 1, row: 1, cell: 1, heading: 'DAVALILAR' } },
        { example: 'İç tablo', location: { kind: 'cell', table: 2, row: 1, cell: 2, heading: 'DAVALILAR' } }
    ]);
});

test('a centred table and the centred table nested in it are each reported once', async () => {
    const nested = table({ tblPr: jc('right'), rows: [[p(r('İç'))]] });
    const { parsed } = await convertBody(table({ tblPr: jc('center'), rows: [[p(r('Dış')) + nested]] }));
    assert.equal(count(parsed, 'tablePlacement'), 2);
    assert.deepEqual(examples(parsed, 'tablePlacement').sort(), ['Dış', 'İç'].sort());
    for (const { location } of occurrences(parsed, 'tablePlacement')) {
        assert.deepEqual(location, { kind: 'cell', table: 1, row: 1, cell: 1 });
    }
});

test('the example is the first non-empty paragraph, and a table with no text has none', async () => {
    const { parsed } = await convertBody(
        table({ tblPr: jc('center'), rows: [[p(''), p(r('   '))], [p(r('İkinci satır')), p(r('x'))]] }) + p(r('-')) +
        table({ tblPr: jc('center'), rows: [[p('')]] }));
    assert.equal(count(parsed, 'tablePlacement'), 2);
    assert.deepEqual(occurrences(parsed, 'tablePlacement'), [
        { example: 'İkinci satır', location: { kind: 'cell', table: 1, row: 1, cell: 1 } },
        { example: '', location: { kind: 'cell', table: 2, row: 1, cell: 1 } }
    ]);
});

test('a table flattened past the nesting limit, one in a cell written empty and one whose rows are all deleted give no note', async () => {
    const centred = text => table({ tblPr: jc('center'), rows: [[p(r(text))]] });
    // Four levels: the innermost is flattened into its cell (nestedTables)
    const deep = table({ rows: [[table({ rows: [[table({ rows: [[centred('Düzleşen')]] })]] })]] });
    const merged = '<w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>' +
        `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p(r('Birleşik'))}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${centred('Yazılmayan')}${p('')}</w:tc></w:tr></w:tbl>`;
    const deleted = '<w:tbl><w:tblPr><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>' +
        `<w:tr><w:trPr><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:trPr><w:tc>${p(r('Silinen'))}</w:tc></w:tr></w:tbl>`;
    const { parsed } = await convertBody(deep + p(r('-')) + merged + p(r('-')) + deleted);
    assert.equal(finding(parsed, 'tablePlacement'), undefined);
    assert.deepEqual(examples(parsed, 'nestedTables'), ['Düzleşen']);
});

test('a centred table in a text box is reported at the paragraph that anchors it', async () => {
    const textBox = inner => '<w:r><w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:graphic>' +
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const { parsed } = await convertBody(p(r('Kutu: ') + textBox(twoColumns(jc('center'), 'Kutudaki tablo') + p(''))));
    assert.deepEqual(occurrences(parsed, 'tablePlacement'),
        [{ example: 'Kutudaki tablo', location: { kind: 'paragraph', paragraph: 1 } }]);
});

test('the petition fixtures give no tablePlacement note', async () => {
    for (const petition of petitions) {
        const { DocxParser } = loadConverter();
        const parsed = await new DocxParser().parse(await petition.build());
        assert.ok(!plain(parsed.findings).some(f => f.code === 'tablePlacement'), petition.name);
    }
});
