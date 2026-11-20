const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, loadConverter, buildDocx } = require('./helpers');

const p = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

// A table on a grid of 2000-twip columns; each cell is [text, tcPr]
const table = (rows, tblPr = '') => {
    const columns = Math.max(...rows.map(cells => cells.reduce((sum, [, tcPr = '']) =>
        sum + Number((/w:gridSpan w:val="(\d+)"/.exec(tcPr) || [0, 1])[1]), 0)));
    return `<w:tbl>${tblPr ? `<w:tblPr>${tblPr}</w:tblPr>` : ''}` +
        `<w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>` +
        rows.map(cells => `<w:tr>${cells.map(([text, tcPr = '']) =>
            `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${p(text)}</w:tc>`).join('')}</w:tr>`).join('') +
        '</w:tbl>';
};

// Border elements: edges('single', 'top', 'insideH') -> <w:top w:val="single" .../><w:insideH .../>
const edges = (val, ...names) => names
    .map(name => `<w:${name} w:val="${val}" w:sz="4" w:space="0" w:color="auto"/>`).join('');
const ALL_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
const tblBorders = inner => `<w:tblBorders>${inner}</w:tblBorders>`;
const tcBorders = inner => `<w:tcBorders>${inner}</w:tcBorders>`;
const style = val => `<w:tblStyle w:val="${val}"/>`;
const VMERGE_RESTART = '<w:vMerge w:val="restart"/>';
const VMERGE_CONTINUE = '<w:vMerge/>';

// Word's default table style ("Normal Table"): cell margins, no borders
const NORMAL_TABLE = '<w:style w:type="table" w:default="1" w:styleId="NormalTable"><w:name w:val="Normal Table"/>' +
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>';
const tableStyle = (id, borders, basedOn = 'NormalTable') =>
    `<w:style w:type="table" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="${basedOn}"/>` +
    `<w:tblPr>${borders ? tblBorders(borders) : ''}</w:tblPr></w:style>`;
// "Table Grid": every edge drawn, through w:basedOn
const TABLE_GRID = tableStyle('TableGrid', edges('single', ...ALL_EDGES));
const STYLES = NORMAL_TABLE + TABLE_GRID +
    // Based on Table Grid, sets nothing itself
    tableStyle('GridChild', '', 'TableGrid') +
    // Horizontal rules only
    tableStyle('Rules', edges('single', 'top', 'bottom', 'insideH'));

const tableElement = doc => doc.querySelector('elements table');
const cellRows = doc => [...doc.querySelectorAll('elements table row')].map(row => [...row.querySelectorAll('cell')]);
const attributes = (doc, name) => cellRows(doc).map(cells => cells.map(cell => cell.getAttribute(name)));

test('a table with no borders and no style takes the default table style: no borders', async () => {
    const { doc } = await convertBody(table([[['A1'], ['B1']], [['A2'], ['B2']]]), { styles: NORMAL_TABLE });

    assert.equal(tableElement(doc).getAttribute('border'), 'borderNone');
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, null]]);
    assert.deepEqual(attributes(doc, 'border'), [[null, null], [null, null]]);
});

test('a styles part with no table style at all gives tables no borders', async () => {
    const { doc } = await convertBody(table([[['A1'], ['B1']]]), { styles: '' });
    assert.equal(tableElement(doc).getAttribute('border'), 'borderNone');
});

test('a style that draws every edge, directly or through w:basedOn, keeps a full grid with no borderSpec', async () => {
    for (const id of ['TableGrid', 'GridChild']) {
        const { doc } = await convertBody(table([[['A1'], ['B1']], [['A2'], ['B2']]], style(id)), { styles: STYLES });

        assert.equal(tableElement(doc).getAttribute('border'), 'borderCell', id);
        assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, null]], id);
        assert.deepEqual(attributes(doc, 'border'), [[null, null], [null, null]], id);
    }
});

test('a table naming a missing style falls back to the default table style', async () => {
    const { doc } = await convertBody(table([[['A1']]], style('Missing')), { styles: STYLES });
    assert.equal(tableElement(doc).getAttribute('border'), 'borderNone');
});

test('horizontal rules only: every cell gets the top and bottom bits, row by row', async () => {
    const rows = [[['A1'], ['B1']], [['A2'], ['B2']], [['A3'], ['B3']]];
    const expected = [['5', '5'], ['5', '5'], ['5', '5']];

    // From a style, and from the table's own w:tblBorders under the default style
    const { doc } = await convertBody(table(rows, style('Rules')), { styles: STYLES });
    assert.equal(tableElement(doc).getAttribute('border'), 'borderCell');
    assert.deepEqual(attributes(doc, 'borderSpec'), expected);

    const own = await convertBody(table(rows, tblBorders(edges('single', 'top', 'bottom', 'insideH'))), { styles: STYLES });
    assert.deepEqual(attributes(own.doc, 'borderSpec'), expected);
});

test("a table's own w:tblBorders removes only the edges it sets to nil, keeping its style's others", async () => {
    const { doc } = await convertBody(
        table([[['A1'], ['B1']], [['A2'], ['B2']]], style('TableGrid') + tblBorders(edges('nil', 'top'))),
        { styles: STYLES });

    assert.equal(tableElement(doc).getAttribute('border'), 'borderCell');
    // The first row loses its top edge (1); the second row keeps all four
    assert.deepEqual(attributes(doc, 'borderSpec'), [['14', '14'], [null, null]]);
});

test('an outer frame only draws each cell edge on the outside of the table', async () => {
    const { doc } = await convertBody(
        table([[['A1'], ['B1']], [['A2'], ['B2']]], tblBorders(edges('single', 'top', 'left', 'bottom', 'right'))),
        { styles: STYLES });

    // top+left, top+right, bottom+left, bottom+right
    assert.deepEqual(attributes(doc, 'borderSpec'), [['9', '3'], ['12', '6']]);
});

test('a cell with w:tcBorders all nil loses its own edges; a neighbour keeps a shared edge it draws', async () => {
    const nil = tcBorders(edges('nil', 'top', 'left', 'bottom', 'right'));
    const { doc } = await convertBody(table([[['A1', nil], ['B1']], [['A2'], ['B2']]], style('TableGrid')), { styles: STYLES });

    // A1 loses its outer top and left edges; B1's left and A2's top still draw the shared ones
    assert.deepEqual(attributes(doc, 'borderSpec'), [['6', null], [null, null]]);
});

test('a shared edge both cells set to nil is drawn on neither', async () => {
    const nil = tcBorders(edges('nil', 'top', 'left', 'bottom', 'right'));
    const { doc } = await convertBody(table([
        [['A1', nil], ['B1', tcBorders(edges('nil', 'left'))]],
        [['A2', tcBorders(edges('nil', 'top'))], ['B2']]
    ], style('TableGrid')), { styles: STYLES });

    assert.equal(tableElement(doc).getAttribute('border'), 'borderCell');
    assert.deepEqual(attributes(doc, 'border'), [['borderNone', null], [null, null]]);
    // B1 without its left edge (8), A2 without its top edge (1)
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, '7'], ['14', null]]);
});

test('a single bottom border on a signature cell in a borderless table writes borderSpec 4 on that cell only', async () => {
    const { doc } = await convertBody(table([
        [[''], ['Davalı']],
        [[''], ['Ad Soyad', tcBorders(edges('single', 'bottom'))]]
    ]), { styles: STYLES });

    assert.equal(tableElement(doc).getAttribute('border'), 'borderCell');
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, '4']]);
    assert.deepEqual(attributes(doc, 'border'), [['borderNone', 'borderNone'], ['borderNone', null]]);
});

test('a cell edge is set on both cells that share it', async () => {
    const { doc } = await convertBody(table([
        [['A1', tcBorders(edges('single', 'bottom', 'right'))], ['B1']],
        [['A2'], ['B2']]
    ]), { styles: STYLES });

    assert.deepEqual(attributes(doc, 'borderSpec'), [['6', '8'], ['1', null]]);
    assert.deepEqual(attributes(doc, 'border'), [[null, null], [null, 'borderNone']]);
});

test("a line along part of a wide cell's edge stays on the narrower cell that sets it", async () => {
    const WIDE = '<w:gridSpan w:val="2"/>';

    // Under the wide cell: only A2 draws the line, so the wide cell's bottom isn't set
    const below = await convertBody(table([
        [['Wide', WIDE]],
        [['A2', tcBorders(edges('single', 'top'))], ['B2']]
    ]), { styles: STYLES });
    assert.deepEqual(attributes(below.doc, 'borderSpec'), [[null], ['1', null]]);
    assert.deepEqual(attributes(below.doc, 'border'), [['borderNone'], [null, 'borderNone']]);

    // Above the wide cell: only A1 draws the line, so the wide cell's top isn't set
    const above = await convertBody(table([
        [['A1', tcBorders(edges('single', 'bottom'))], ['B1']],
        [['Wide', WIDE]]
    ]), { styles: STYLES });
    assert.deepEqual(attributes(above.doc, 'borderSpec'), [['4', null], [null]]);
    assert.deepEqual(attributes(above.doc, 'border'), [[null, 'borderNone'], ['borderNone']]);

    // A line along the wide cell's whole edge is still set on the cells above it
    const whole = await convertBody(table([
        [['A1'], ['B1']],
        [['Wide', WIDE + tcBorders(edges('single', 'top'))]]
    ]), { styles: STYLES });
    assert.deepEqual(attributes(whole.doc, 'borderSpec'), [['4', '4'], ['1']]);
});

test('start and end are read like left and right, in w:tblBorders and w:tcBorders', async () => {
    const { doc } = await convertBody(table([
        [['A1'], ['B1'], ['C1', tcBorders(edges('single', 'start'))]]
    ], tblBorders(edges('single', 'start', 'end'))), { styles: STYLES });

    // Left of A1, right of C1, and the edge between B1 and C1 from C1's start
    assert.deepEqual(attributes(doc, 'borderSpec'), [['8', '2', '10']]);

    const removed = await convertBody(table([[['A1', tcBorders(edges('nil', 'end'))], ['B1']]],
        style('TableGrid') + tblBorders(edges('nil', 'start'))), { styles: STYLES });
    // A1 loses its left (table start) edge; its end edge stays because B1's left draws it
    assert.deepEqual(attributes(removed.doc, 'borderSpec'), [['7', null]]);
});

test('a vertical merge in a partly bordered table joins its cells within the resolved edges', async () => {
    const { doc } = await convertBody(table([
        [['Merged', VMERGE_RESTART], ['B1']],
        [['', VMERGE_CONTINUE], ['B2']],
        [['A3'], ['B3']]
    ], style('Rules')), { styles: STYLES });

    // Top and bottom only (5); the merge drops its bottom (4) from the first cell and its top (1) from the last
    assert.deepEqual(attributes(doc, 'borderSpec'), [['1', '5'], ['4', '5'], ['5', '5']]);
});

test('a vertical merge in a borderless table writes no border attributes on its cells', async () => {
    const { doc } = await convertBody(table([
        [['Merged', VMERGE_RESTART], ['B1']],
        [['', VMERGE_CONTINUE], ['B2']]
    ]), { styles: STYLES });

    assert.equal(tableElement(doc).getAttribute('border'), 'borderNone');
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, null]]);
    assert.deepEqual(attributes(doc, 'border'), [[null, null], [null, null]]);
});

test('without a styles part, a table with no borders of its own keeps a full grid', async () => {
    const { doc, parsed } = await convertBody(table([[['A1'], ['B1']]]));
    assert.equal(parsed.elements[0].border, 'borderCell');
    assert.equal(tableElement(doc).getAttribute('border'), 'borderCell');
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null]]);

    // Its own nil edge still removes just that edge
    const own = await convertBody(table([[['A1'], ['B1']]], tblBorders(edges('nil', 'top'))));
    assert.deepEqual(attributes(own.doc, 'borderSpec'), [['14', '14']]);
});

test("a table style's cell borders (w:tcPr/w:tcBorders) apply to every cell, under the cell's own", async () => {
    const cellStyle = (id, borders, basedOn = 'NormalTable') =>
        `<w:style w:type="table" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="${basedOn}"/>` +
        `<w:tblPr/><w:tcPr>${tcBorders(borders)}</w:tcPr></w:style>`;
    const styles = STYLES + cellStyle('CellGrid', edges('single', 'top', 'left', 'bottom', 'right')) +
        // Through w:basedOn, and removing one cell edge over the table's
        tableStyle('CellGridChild', '', 'CellGrid') + cellStyle('GridNoTop', edges('nil', 'top'), 'TableGrid');
    const rows = [[['A1'], ['B1']], [['A2'], ['B2']]];

    for (const id of ['CellGrid', 'CellGridChild']) {
        const { doc } = await convertBody(table(rows, style(id)), { styles });
        assert.equal(tableElement(doc).getAttribute('border'), 'borderCell', id);
        assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, null]], id);
    }

    // Every cell drops its top; the edge between the rows stays because the upper cells draw it
    const noTop = await convertBody(table(rows, style('GridNoTop')), { styles });
    assert.deepEqual(attributes(noTop.doc, 'borderSpec'), [['14', '14'], [null, null]]);

    // A cell's own w:tcBorders overrides the style's cell borders
    const own = await convertBody(table([[['A1', tcBorders(edges('nil', 'top', 'left'))], ['B1']]], style('CellGrid')), { styles });
    assert.deepEqual(attributes(own.doc, 'borderSpec'), [['6', null]]);
});

test('a w:basedOn loop between table styles stops instead of hanging', async () => {
    const styles = tableStyle('Loop1', edges('single', 'top'), 'Loop2') + tableStyle('Loop2', '', 'Loop1');
    const { doc } = await convertBody(table([[['A1']]], style('Loop1')), { styles });
    assert.deepEqual(attributes(doc, 'borderSpec'), [['1']]);
});

test('parsed cells carry their resolved edges', async () => {
    const { DocxParser } = loadConverter();
    const parsed = await new DocxParser().parse(await buildDocx(
        table([[['A1'], ['B1']]], style('Rules')), { styles: STYLES }));
    const [a1, b1] = parsed.elements[0].rows[0].cells;
    assert.deepEqual({ ...a1.borders }, { top: true, right: false, bottom: true, left: false });
    assert.deepEqual({ ...b1.borders }, { top: true, right: false, bottom: true, left: false });
});

test('a reused parser does not carry table styles into a document without a styles part', async () => {
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();
    const first = await parser.parse(await buildDocx(table([[['A1']]]), { styles: NORMAL_TABLE }));
    assert.equal(first.elements[0].border, 'borderNone');
    const second = await parser.parse(await buildDocx(table([[['A1']]])));
    assert.equal(second.elements[0].border, 'borderCell');
});
