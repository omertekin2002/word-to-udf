/**
 * Tables nested in table cells: written as a <table> inside the <cell>, which the documented
 * format allows ("Hücreler paragraflar ve iç içe tablolar içerebilir"), down to MAX_TABLE_DEPTH
 * levels of tables. A table nested deeper is flattened into its cell and reported.
 */

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
// A table: tblPr, grid column widths in twips, and rows of cells, each cell [content, tcPr]
const table = ({ tblPr = '', grid = [4000], rows }) => '<w:tbl>' +
    (tblPr ? `<w:tblPr>${tblPr}</w:tblPr>` : '') +
    `<w:tblGrid>${grid.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    rows.map(cells => `<w:tr>${cells.map(([content, tcPr = '']) =>
        `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${content}</w:tc>`).join('')}</w:tr>`).join('') +
    '</w:tbl>';
// A one-cell table holding content
const box = content => table({ rows: [[[content]]] });
const borders = (tag, val, edges, { sz = 4, color = 'auto' } = {}) =>
    `<w:${tag}>${edges.map(edge => `<w:${edge} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`).join('')}</w:${tag}>`;
const ALL_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
// A floating text box the way Word writes it, DrawingML with a VML fallback
const textBox = inner => '<w:r><mc:AlternateContent>' +
    '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';

const occurrences = (text, needle) => text.split(needle).length - 1;
const children = el => [...el.children].map(child => child.tagName);
// The text of each <content> directly in an element's paragraphs, not in its nested tables
const ownTexts = (el, content) => [...el.children].filter(child => child.tagName === 'paragraph')
    .map(para => [...para.children].map(run => textOf(run, content)).join(''));
const plain = value => JSON.parse(JSON.stringify(value));

// The petition case: a layout table whose cell holds a label, a 2×2 amount schedule, then a note
const SCHEDULE = table({
    tblPr: borders('tblBorders', 'nil', ALL_EDGES),
    grid: [6000, 3000],
    rows: [[[p(r('Davalılar:')) + table({
        grid: [3000, 1500],
        rows: [
            [[p(r('Asıl alacak'))], [p(r('10.000 TL'))]],
            [[p(r('Faiz'))], [p(r('1.250 TL'))]]
        ]
    }) + p(r('Toplam yukarıdadır.'))], [p(r('Sağ hücre'))]]]
});

test('a 2-level table writes one <table> under <elements> and one inside the cell, in document order', async () => {
    const { doc, content, parsed } = await convertBody(SCHEDULE);

    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.equal(doc.querySelectorAll('elements > table').length, 1);
    const [outerCell, rightCell] = doc.querySelectorAll('elements > table > row > cell');
    assert.deepEqual(children(outerCell), ['paragraph', 'table', 'paragraph']);
    assert.deepEqual(children(rightCell), ['paragraph']);
    assert.deepEqual(ownTexts(outerCell, content), ['Davalılar:', 'Toplam yukarıdadır.']);

    const nested = outerCell.querySelector(':scope > table');
    const rows = [...nested.querySelectorAll(':scope > row')];
    assert.deepEqual(rows.map(row => [...row.querySelectorAll(':scope > cell')].map(cell => ownTexts(cell, content))),
        [[['Asıl alacak'], ['10.000 TL']], [['Faiz'], ['1.250 TL']]]);
    assert.deepEqual(rows.map(row => row.getAttribute('rowName')), ['row1', 'row2']);
});

test('each nested cell\'s text appears once, in reading order, and every offset points at its own text', async () => {
    const { doc, content } = await convertBody(SCHEDULE);
    const texts = ['Davalılar:', 'Asıl alacak', '10.000 TL', 'Faiz', '1.250 TL', 'Toplam yukarıdadır.', 'Sağ hücre'];

    for (const text of texts) assert.equal(occurrences(content, text), 1, text);
    const order = texts.map(text => content.indexOf(text));
    assert.ok(order.every((at, i) => i === 0 || at > order[i - 1]), JSON.stringify(content));
    // Every run, in document order, covers the next piece of the content, with nothing written twice
    const runs = [...doc.querySelectorAll('elements content[startOffset]')];
    assert.deepEqual(runs.map(run => textOf(run, content)).filter(text => text.trim() !== ''), texts);
    let end = 0;
    for (const run of runs) {
        const start = Number(run.getAttribute('startOffset'));
        assert.ok(start >= end, `overlapping run at ${start}`);
        end = start + Number(run.getAttribute('length'));
    }
    // The label paragraph is followed by a line break, as between a cell's paragraphs; the table isn't
    assert.equal(content, 'Davalılar:\nAsıl alacak10.000 TLFaiz1.250 TLToplam yukarıdadır.Sağ hücre');
});

test('a nested table keeps its own widths, spans, merges, shading, alignment and borders, not the outer table\'s', async () => {
    const red = { sz: 12, color: 'FF0000' };
    const inner = table({
        tblPr: borders('tblBorders', 'double', ALL_EDGES, red),
        grid: [1000, 2000, 3000],
        rows: [
            [[p(r('Geniş')), '<w:gridSpan w:val="2"/>'],
                [p(r('Birleşik')), '<w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>']],
            [[p(r('Orta')), '<w:vAlign w:val="center"/>'], [p(r('Düz'))], [p(r('Yazılmaz')), '<w:vMerge/>']]
        ]
    });
    const outer = table({
        tblPr: borders('tblBorders', 'nil', ALL_EDGES),
        grid: [4000, 4000],
        rows: [[[p(r('Dış')) + inner + p(''), '<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>'], [p(r('Yan'))]]]
    });
    const { doc, content } = await convertBody(outer);

    const outerTable = doc.querySelector('elements > table');
    assert.equal(outerTable.getAttribute('columnCount'), '2');
    assert.equal(outerTable.getAttribute('columnSpans'), '150,150');
    assert.equal(outerTable.getAttribute('border'), 'borderNone');
    const outerCell = outerTable.querySelector(':scope > row > cell');
    assert.equal(outerCell.getAttribute('fillColor'), String(0xFF00FF00 | 0));

    const nested = outerCell.querySelector(':scope > table');
    assert.equal(nested.getAttribute('columnCount'), '3');
    assert.equal(nested.getAttribute('columnSpans'), '50,100,150');
    assert.equal(nested.getAttribute('border'), 'borderCell');
    const [[wide, merged], [centred, plainCell, continued]] = [...nested.querySelectorAll(':scope > row')]
        .map(row => [...row.querySelectorAll(':scope > cell')]);
    assert.equal(wide.getAttribute('colspan'), '2');
    // A merged area is shaded the way its first cell is
    assert.equal(merged.getAttribute('fillColor'), String(0xFFFFFF00 | 0));
    assert.equal(continued.getAttribute('fillColor'), String(0xFFFFFF00 | 0));
    assert.equal(centred.getAttribute('align'), 'vcenter');
    assert.equal(plainCell.getAttribute('align'), null);
    // The vertical merge: no bottom edge on its first cell, no top edge on the continuation
    assert.equal(merged.getAttribute('borderSpec'), '11');
    assert.equal(continued.getAttribute('borderSpec'), '14');
    for (const cell of [wide, merged, centred, plainCell, continued]) {
        assert.deepEqual(['borderStyle', 'borderWidth', 'borderColor'].map(name => cell.getAttribute(name)),
            ['borderStyle-double', '1.5', String(0xFFFF0000 | 0)]);
    }
    // The continuation is written empty and adds nothing to the content
    assert.equal(occurrences(content, 'Yazılmaz'), 0);
    assert.deepEqual([...continued.querySelectorAll('content')].map(run => run.getAttribute('length')), ['0']);
});

test('a nested continuation cell\'s content is neither written nor reported', async () => {
    const exact = text => p(r(text), '<w:spacing w:line="240" w:lineRule="exact"/>');
    const inner = table({
        rows: [[[p(r('Başlık')), '<w:vMerge w:val="restart"/>']], [[exact('Gizli devam'), '<w:vMerge/>']]]
    });
    const { content, parsed } = await convertBody(box(p(r('Dış')) + inner + exact('Sonra')));

    assert.ok(!content.includes('Gizli devam'), JSON.stringify(content));
    // Only the paragraph after the nested table, which is written, is reported
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'lineSpacing', count: 1, examples: ['Sonra'] }]);
});

test('a table nested in a text box in a cell is written in that cell, after the anchoring paragraph', async () => {
    const { doc, content, parsed } = await convertBody(
        box(p(r('Çapa') + textBox(p(r('Kutu')) + box(p(r('Kutudaki tablo'))) + p('')))));

    assert.deepEqual(findingSummaries(parsed.findings), []);
    const cell = doc.querySelector('elements > table > row > cell');
    assert.deepEqual(children(cell), ['paragraph', 'paragraph', 'table', 'paragraph']);
    assert.equal(occurrences(content, 'Kutudaki tablo'), 1);
    assert.deepEqual(ownTexts(cell.querySelector(':scope > table > row > cell'), content), ['Kutudaki tablo']);
});

test('three levels of tables are written as tables; a fourth is flattened into its cell and reported once', async () => {
    const three = await convertBody(box(p(r('Bir')) + box(p(r('İki')) + box(p(r('Üç'))) + p('')) + p('')));
    assert.deepEqual(findingSummaries(three.parsed.findings), []);
    assert.equal(three.doc.querySelectorAll('elements table').length, 3);
    assert.equal(three.doc.querySelectorAll('elements > table > row > cell > table > row > cell > table').length, 1);

    const body = p(r('Giriş')) + box(p(r('Bir')) + box(p(r('İki')) + box(p(r('Üç')) + box(p(r('Dört'))) + p(r('Son'))) + p('')) + p(''));
    const four = await convertBody(body);
    assert.equal(four.doc.querySelectorAll('elements table').length, 3);
    assert.equal(occurrences(four.content, 'Dört'), 1);
    const deepest = four.doc.querySelector('elements > table table table > row > cell');
    assert.deepEqual(ownTexts(deepest, four.content), ['Üç', 'Dört', 'Son']);
    assert.deepEqual(plain(four.parsed.findings), [{
        code: 'nestedTables', count: 1, examples: ['Dört'],
        occurrences: [{ example: 'Dört', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }]
    }]);
});

test('a nested table in a body text box\'s table is nested too', async () => {
    // The text box's table is a table of its own (1 level), so a table in its cell is the 2nd
    const { doc, parsed } = await convertBody(p(r('Çapa') + textBox(box(p(r('Kutu')) + box(p(r('İç'))) + p('')) + p(''))));
    assert.deepEqual(findingSummaries(parsed.findings), []);
    assert.equal(doc.querySelectorAll('elements > table > row > cell > table').length, 1);
});
