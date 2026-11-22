/**
 * Table border lines: Word's border style (w:val), width (w:sz) and color (w:color) become the
 * documented <cell borderStyle borderWidth borderColor>, one line per cell (the heaviest it
 * draws), written only when it differs from Word's default line. Which edges are drawn
 * (borderSpec) is covered by tests/borders.test.js.
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { convertBody } = require('./helpers');

const utils = vm.createContext({});
vm.runInContext(fs.readFileSync(path.resolve(__dirname, '..', 'js', 'utils.js'), 'utf8'), utils,
    { filename: 'js/utils.js' });
const { cellBorderLine, heavierBorderLine, isDefaultBorderLine } = utils;

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

// Border elements with a line: line('double', { sz: 6, color: 'FF0000' }, 'top', 'left').
// An attribute given as null is left out.
const line = (val, { sz = 4, color = 'auto' } = {}, ...names) => names.map(name =>
    `<w:${name} w:val="${val}"${sz === null ? '' : ` w:sz="${sz}"`} w:space="0"${color === null ? '' : ` w:color="${color}"`}/>`).join('');
const OUTER = ['top', 'left', 'bottom', 'right'];
const ALL_EDGES = [...OUTER, 'insideH', 'insideV'];
const NO_EDGES = line('nil', {}, ...ALL_EDGES);
const tblBorders = inner => `<w:tblBorders>${inner}</w:tblBorders>`;
const tcBorders = inner => `<w:tcBorders>${inner}</w:tcBorders>`;

// Word's default table style ("Normal Table"): no borders
const NORMAL_TABLE = '<w:style w:type="table" w:default="1" w:styleId="NormalTable"><w:name w:val="Normal Table"/></w:style>';
const tableStyle = (id, borders, basedOn = 'NormalTable') =>
    `<w:style w:type="table" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="${basedOn}"/>` +
    `<w:tblPr>${borders ? tblBorders(borders) : ''}</w:tblPr></w:style>`;

// One single-cell table per [text, tcPr], each drawing no edge but the cell's own, with a
// paragraph between them
const cellTables = cells => cells.map(([text, tcPr]) => table([[[text, tcPr]]], tblBorders(NO_EDGES))).join(p(''));

const BLACK = '-16777216';
const LINE_ATTRIBUTES = ['borderStyle', 'borderWidth', 'borderColor'];
const cellRows = doc => [...doc.querySelectorAll('elements table row')].map(row => [...row.querySelectorAll('cell')]);
// Each cell's [borderStyle, borderWidth, borderColor], or null when it has none of them
const lines = doc => cellRows(doc).map(cells => cells.map(cell => {
    const values = LINE_ATTRIBUTES.map(name => cell.getAttribute(name));
    assert.ok(values.every(v => v === null) || values.every(v => v !== null), 'all three or none');
    return values[0] === null ? null : values;
}));
// The one cell of each table's [borderStyle, borderWidth, borderColor] (see cellTables)
const tableLines = doc => lines(doc).map(([values]) => values);
const attributes = (doc, name) => cellRows(doc).map(cells => cells.map(cell => cell.getAttribute(name)));
const borderFinding = parsed => JSON.parse(JSON.stringify(parsed.findings)).find(f => f.code === 'tableBorders');

// The mapping helper

test('Word border styles map to the four documented ones; anything else is approximated as solid', () => {
    const expected = {
        single: 'borderStyle-solid',
        thick: 'borderStyle-solid',
        dotted: 'borderStyle-dotted',
        dashed: 'borderStyle-dashed',
        dashSmallGap: 'borderStyle-dashed',
        dotDash: 'borderStyle-dashed',
        dotDotDash: 'borderStyle-dashed',
        dashDotStroked: 'borderStyle-dashed',
        double: 'borderStyle-double'
    };
    for (const [val, style] of Object.entries(expected)) {
        const written = cellBorderLine({ val, size: 4, color: 'auto' });
        assert.equal(written.borderStyle, style, val);
        assert.equal(written.documented, true, val);
    }
    for (const val of ['triple', 'thinThickSmallGap', 'thickThinLargeGap', 'wave', 'doubleWave',
        'threeDEmboss', 'threeDEngrave', 'outset', 'inset', 'apples']) {
        const written = cellBorderLine({ val, size: 4, color: 'auto' });
        assert.equal(written.borderStyle, 'borderStyle-solid', val);
        assert.equal(written.documented, false, val);
    }
});

test('widths are eighths of a point, written in points with at least one decimal place', () => {
    const widths = { 2: '0.25', 4: '0.5', 6: '0.75', 8: '1.0', 12: '1.5', 18: '2.25', 24: '3.0', 1: '0.125', 96: '12.0' };
    for (const [size, width] of Object.entries(widths)) {
        assert.equal(cellBorderLine({ val: 'single', size: Number(size), color: 'auto' }).borderWidth, width, size);
    }
});

test('colors become signed ARGB with full alpha; auto is black', () => {
    const color = hex => cellBorderLine({ val: 'single', size: 4, color: hex }).borderColor;
    assert.equal(color('auto'), -16777216);
    assert.equal(color('000000'), -16777216);
    assert.equal(color('FF0000'), -65536);
    assert.equal(color('808080'), -8355712);
    assert.equal(color('FFFFFF'), -1);
});

test('the heavier line is the wider one, then double > solid > dashed > dotted', () => {
    const l = (val, size = 4) => ({ val, size, color: 'auto' });
    assert.equal(heavierBorderLine(null, null), null);
    assert.equal(heavierBorderLine(undefined, null), null);
    const single = l('single');
    assert.equal(heavierBorderLine(single, null), single);
    assert.equal(heavierBorderLine(null, single), single);
    const wideDotted = l('dotted', 12);
    assert.equal(heavierBorderLine(single, wideDotted), wideDotted);
    assert.equal(heavierBorderLine(wideDotted, l('double', 8)), wideDotted);
    const order = ['dotted', 'dashed', 'single', 'double'].map(val => l(val));
    for (let i = 0; i < order.length; i++) {
        for (let j = 0; j < order.length; j++) {
            assert.equal(heavierBorderLine(order[i], order[j]), order[Math.max(i, j)], `${order[i].val} vs ${order[j].val}`);
        }
    }
    // An approximated style weighs as solid; on a tie the first line is kept
    const wave = l('wave');
    assert.equal(heavierBorderLine(wave, single), wave);
    assert.equal(heavierBorderLine(single, wave), single);
});

test("Word's default line is single, 4 eighths, automatic or black", () => {
    assert.equal(isDefaultBorderLine({ val: 'single', size: 4, color: 'auto' }), true);
    assert.equal(isDefaultBorderLine({ val: 'single', size: 4, color: '000000' }), true);
    assert.equal(isDefaultBorderLine({ val: 'single', size: 6, color: 'auto' }), false);
    assert.equal(isDefaultBorderLine({ val: 'single', size: 4, color: '000001' }), false);
    assert.equal(isDefaultBorderLine({ val: 'thick', size: 4, color: 'auto' }), false);
    assert.equal(isDefaultBorderLine({ val: 'wave', size: 4, color: 'auto' }), false);
});

// Conversion

test('a double red w:tblBorders writes its line on every cell, with no note', async () => {
    const { doc, parsed } = await convertBody(table([[['A1'], ['B1']], [['A2'], ['B2']]],
        tblBorders(line('double', { sz: 6, color: 'FF0000' }, ...ALL_EDGES))), { styles: NORMAL_TABLE });
    const red = ['borderStyle-double', '0.75', '-65536'];
    assert.deepEqual(lines(doc), [[red, red], [red, red]]);
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, null], [null, null]]);
    assert.equal(borderFinding(parsed), undefined);
});

test('dotted, dashed and the dash variants map to borderStyle-dotted and borderStyle-dashed', async () => {
    const vals = ['dotted', 'dashed', 'dashSmallGap', 'dotDash', 'dotDotDash', 'dashDotStroked'];
    const { doc, parsed } = await convertBody(cellTables(vals.map(val => [val, tcBorders(line(val, {}, ...OUTER))])),
        { styles: NORMAL_TABLE });
    assert.deepEqual(tableLines(doc), [['borderStyle-dotted', '0.5', BLACK], ...Array(5).fill(['borderStyle-dashed', '0.5', BLACK])]);
    assert.equal(borderFinding(parsed), undefined);
});

test('w:sz 12 is 1.5 pt; a missing or unreadable w:sz is 4 (½ pt)', async () => {
    const { doc } = await convertBody(cellTables([
        ['Wide', tcBorders(line('double', { sz: 12 }, ...OUTER))],
        ['No size', tcBorders(line('double', { sz: null }, ...OUTER))],
        ['Bad size', tcBorders(line('double', { sz: 'kalın' }, ...OUTER))],
        ['Zero size', tcBorders(line('double', { sz: 0 }, ...OUTER))]
    ]), { styles: NORMAL_TABLE });
    assert.deepEqual(tableLines(doc).map(values => values[1]), ['1.5', '0.5', '0.5', '0.5']);
});

test('auto, a missing and an unreadable w:color are black', async () => {
    const { doc } = await convertBody(cellTables([
        ['Auto', tcBorders(line('dashed', { color: 'auto' }, ...OUTER))],
        ['None', tcBorders(line('dashed', { color: null }, ...OUTER))],
        ['Bad', tcBorders(line('dashed', { color: 'kırmızı' }, ...OUTER))],
        ['Lower case', tcBorders(line('dashed', { color: '00ff00' }, ...OUTER))]
    ]), { styles: NORMAL_TABLE });
    assert.deepEqual(tableLines(doc).map(values => values[2]), [BLACK, BLACK, BLACK, '-16711936']);
});

test('default single ½ pt auto or black lines write none of the three attributes', async () => {
    const grid = tableStyle('TableGrid', line('single', {}, ...ALL_EDGES));
    const black = tableStyle('BlackGrid', line('single', { color: '000000' }, ...ALL_EDGES));
    for (const [styles, tblPr] of [
        [NORMAL_TABLE + grid, '<w:tblStyle w:val="TableGrid"/>'],
        [NORMAL_TABLE + black, '<w:tblStyle w:val="BlackGrid"/>'],
        [undefined, ''] // No styles part: a full grid of default lines
    ]) {
        const { doc } = await convertBody(table([[['A1'], ['B1']], [['A2'], ['B2']]], tblPr), { styles });
        assert.equal(doc.querySelector('elements table').getAttribute('border'), 'borderCell');
        assert.deepEqual(lines(doc), [[null, null], [null, null]]);
    }
});

test("a line from the table style chain is inherited, and a cell's own w:tcBorders overrides it for that cell only", async () => {
    // Blue dashed 1.5 pt outer frame and inside rows from a style based on another; no insideV,
    // so the two cells share no edge
    const styles = NORMAL_TABLE + tableStyle('Base', line('dashed', { sz: 12, color: '0000FF' }, ...OUTER, 'insideH')) +
        tableStyle('Child', '', 'Base');
    const own = tcBorders(line('double', {}, 'top', 'left', 'bottom') + line('nil', {}, 'right'));
    const { doc, parsed } = await convertBody(table([[['Own', own], ['Inherited']]], '<w:tblStyle w:val="Child"/>'), { styles });
    assert.deepEqual(lines(doc), [[['borderStyle-double', '0.5', BLACK], ['borderStyle-dashed', '1.5', '-16776961']]]);
    assert.deepEqual(attributes(doc, 'borderSpec'), [['13', '7']]);
    assert.equal(borderFinding(parsed), undefined);
});

test('two cells sharing an edge with different lines both get the heavier line', async () => {
    const { doc } = await convertBody(
        table([[['A', tcBorders(line('single', {}, 'right'))], ['B', tcBorders(line('double', { sz: 12, color: 'FF0000' }, 'left'))]]],
            tblBorders(NO_EDGES)) +
        p('Equal width') +
        table([[['C', tcBorders(line('dashed', { sz: 8 }, 'right'))], ['D', tcBorders(line('dotted', { sz: 8 }, 'left'))]]],
            tblBorders(NO_EDGES)) +
        p('Above and below') +
        table([[['E', tcBorders(line('dotted', {}, 'bottom'))]], [['F', tcBorders(line('double', {}, 'top'))]]],
            tblBorders(NO_EDGES)),
        { styles: NORMAL_TABLE });

    const tables = [...doc.querySelectorAll('elements table')].map(tableEl =>
        [...tableEl.querySelectorAll('cell')].map(cell => [cell.getAttribute('borderSpec'), ...LINE_ATTRIBUTES.map(n => cell.getAttribute(n))]));
    assert.deepEqual(tables[0], [['2', 'borderStyle-double', '1.5', '-65536'], ['8', 'borderStyle-double', '1.5', '-65536']]);
    assert.deepEqual(tables[1], [['2', 'borderStyle-dashed', '1.0', BLACK], ['8', 'borderStyle-dashed', '1.0', BLACK]]);
    assert.deepEqual(tables[2], [['4', 'borderStyle-double', '0.5', BLACK], ['1', 'borderStyle-double', '0.5', BLACK]]);
});

test('a cell whose edges draw different lines gets its heaviest and one located note', async () => {
    // The cell's single right edge is shared with its neighbour, which draws only that one line
    const mixed = tcBorders(line('single', {}, 'top', 'left', 'right') + line('double', { sz: 6 }, 'bottom'));
    const { doc, parsed } = await convertBody(
        p('Önce') + table([[['Taraflar', mixed], ['Komşu']]], tblBorders(NO_EDGES)),
        { styles: NORMAL_TABLE });
    assert.deepEqual(lines(doc), [[['borderStyle-double', '0.75', BLACK], null]]);
    assert.deepEqual(attributes(doc, 'borderSpec'), [[null, '8']]);
    assert.deepEqual(borderFinding(parsed), {
        code: 'tableBorders', count: 1, examples: ['Taraflar'],
        occurrences: [{ example: 'Taraflar', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }]
    });
});

test('a wave or triple edge is written as borderStyle-solid and reported once per cell', async () => {
    const { doc, parsed } = await convertBody(cellTables([
        ['Dalga', tcBorders(line('wave', {}, ...OUTER))],
        ['Üçlü', tcBorders(line('triple', { sz: 12 }, ...OUTER))]
    ]), { styles: NORMAL_TABLE });
    assert.deepEqual(tableLines(doc), [['borderStyle-solid', '0.5', BLACK], ['borderStyle-solid', '1.5', BLACK]]);
    const finding = borderFinding(parsed);
    assert.equal(finding.count, 2);
    assert.deepEqual(finding.examples, ['Dalga', 'Üçlü']);
    assert.deepEqual(finding.occurrences.map(({ location }) => location.table), [1, 2]);
});

test('borderless tables and cells without edges get none of the three attributes', async () => {
    const { doc } = await convertBody(
        table([[['A'], ['B']]], tblBorders(NO_EDGES)) +
        p('Between') +
        table([[['Top', tcBorders(line('double', { color: 'FF0000' }, 'top'))], ['None']]], tblBorders(NO_EDGES)),
        { styles: NORMAL_TABLE });
    const [borderless, partly] = doc.querySelectorAll('elements table');
    assert.equal(borderless.getAttribute('border'), 'borderNone');
    for (const cell of borderless.querySelectorAll('cell')) {
        for (const name of LINE_ATTRIBUTES) assert.equal(cell.getAttribute(name), null);
    }
    const [top, none] = partly.querySelectorAll('cell');
    assert.equal(top.getAttribute('borderStyle'), 'borderStyle-double');
    assert.equal(none.getAttribute('border'), 'borderNone');
    for (const name of LINE_ATTRIBUTES) assert.equal(none.getAttribute(name), null);
});

test('a vertical merge leaves out the edges it joins when choosing the line and deciding the note', async () => {
    // A double frame with a single line between the rows: merged, the two cells draw only the
    // frame; unmerged, each also draws the single line, and is reported
    const tblPr = tblBorders(line('double', { sz: 12 }, ...OUTER) + line('single', {}, 'insideH', 'insideV'));
    const double = ['borderStyle-double', '1.5', BLACK];

    const merged = await convertBody(table([[['Merged', '<w:vMerge w:val="restart"/>']], [['', '<w:vMerge/>']]], tblPr),
        { styles: NORMAL_TABLE });
    assert.deepEqual(lines(merged.doc), [[double], [double]]);
    assert.deepEqual(attributes(merged.doc, 'borderSpec'), [['11'], ['14']]);
    assert.equal(borderFinding(merged.parsed), undefined);

    const unmerged = await convertBody(table([[['Top']], [['Bottom']]], tblPr), { styles: NORMAL_TABLE });
    assert.deepEqual(lines(unmerged.doc), [[double], [double]]);
    assert.deepEqual(borderFinding(unmerged.parsed).occurrences.map(({ location }) => location.row), [1, 2]);
});

test('a nested table\'s cells give the border note like a body table\'s, at the outer cell', async () => {
    const inner = table([[['Inner', tcBorders(line('wave', {}, ...OUTER))]]]);
    const { parsed, doc } = await convertBody(
        `<w:tbl><w:tblPr>${tblBorders(NO_EDGES)}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc>${p('Outer')}${inner}${p('')}</w:tc></w:tr></w:tbl>`,
        { styles: NORMAL_TABLE });
    assert.deepEqual(borderFinding(parsed).occurrences,
        [{ example: 'Inner', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }]);
    // The wave is written as solid on the nested cell, which draws its own four edges
    const nestedCell = doc.querySelector('elements > table cell table cell');
    assert.equal(nestedCell.getAttribute('borderStyle'), 'borderStyle-solid');
    assert.equal(doc.querySelector('elements > table').getAttribute('border'), 'borderNone');
});

test('a table nested past the depth limit, flattened to paragraphs, gives no border note', async () => {
    const inner = table([[['Inner', tcBorders(line('wave', {}, ...OUTER))]]]);
    const wrap = content => `<w:tbl><w:tblPr>${tblBorders(NO_EDGES)}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc>${content}${p('')}</w:tc></w:tr></w:tbl>`;
    const { parsed } = await convertBody(wrap(wrap(wrap(inner))), { styles: NORMAL_TABLE });
    assert.equal(borderFinding(parsed), undefined);
});
