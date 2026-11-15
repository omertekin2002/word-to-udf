const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const run = text => `<w:r><w:t>${text}</w:t></w:r>`;
const tabs = (...stops) => `<w:tabs>${stops.join('')}</w:tabs>`;
const tab = (val, pos, leader) => `<w:tab w:val="${val}" w:pos="${pos}"${leader ? ` w:leader="${leader}"` : ''}/>`;
const styled = (styleId, runs, pPr = '') => p(runs, `<w:pStyle w:val="${styleId}"/>${pPr}`);
const paragraphStyle = (id, { basedOn, pPr = '' } = {}) =>
    `<w:style w:type="paragraph" w:styleId="${id}">` +
    (basedOn ? `<w:basedOn w:val="${basedOn}"/>` : '') +
    (pPr ? `<w:pPr>${pPr}</w:pPr>` : '') + '</w:style>';

/** The TabSet attribute of the paragraph holding exactly this text (null when absent) */
function tabSetOf(doc, content, text) {
    const found = [...doc.querySelectorAll('elements content[startOffset]')].filter(el => textOf(el, content) === text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} elements`);
    return found[0].closest('paragraph').getAttribute('TabSet');
}

test('left stops become TabSet positions in points, to one decimal', async () => {
    const { doc, content } = await convertBody(
        p(run('Party'), tabs(tab('left', 1985), tab('left', 2268))));

    // 1985 / 20 = 99.25 -> 99.3; 2268 / 20 = 113.4
    assert.equal(tabSetOf(doc, content, 'Party'), '99.3:0:0,113.4:0:0');
});

test('stop alignments and leaders map to the documented codes', async () => {
    const cases = [
        [tab('right', 1440, 'dot'), '72.0:2:1'],
        [tab('center', 1440), '72.0:1:0'],
        [tab('decimal', 1440, 'hyphen'), '72.0:3:2'],
        [tab('end', 1440), '72.0:2:0'],
        [tab('start', 1440), '72.0:0:0'],
        [tab('num', 1440), '72.0:0:0'],
        [tab('left', 1440, 'middleDot'), '72.0:0:1'],
        [tab('left', 1440, 'underscore'), '72.0:0:0'],
        [tab('left', 1440, 'heavy'), '72.0:0:0'],
        [tab('left', 1440, 'none'), '72.0:0:0']
    ];
    const { doc, content } = await convertBody(
        cases.map(([stop], i) => p(run(`Case ${i}`), tabs(stop))).join(''));

    cases.forEach(([stop, expected], i) => assert.equal(tabSetOf(doc, content, `Case ${i}`), expected, stop));
});

test('stops from a paragraph style and its w:basedOn parent both apply', async () => {
    const { doc, content, parsed } = await convertBody(
        styled('Child', run('Styled')) + p(run('Unstyled')), {
            styles: paragraphStyle('Base', { pPr: tabs(tab('left', 720)) }) +
                paragraphStyle('Child', { basedOn: 'Base', pPr: tabs(tab('right', 9000, 'dot')) })
        });

    assert.equal(tabSetOf(doc, content, 'Styled'), '36.0:0:0,450.0:2:1');
    assert.equal(tabSetOf(doc, content, 'Unstyled'), null);
    assert.deepEqual(JSON.parse(JSON.stringify(parsed.elements[0].tabStops)), [
        { position: 36, alignment: 0, leader: 0 },
        { position: 450, alignment: 2, leader: 1 }
    ]);
    assert.equal(parsed.elements[0].tabs, undefined, 'the raw tab operations are not exposed');
});

test('stops from w:docDefaults apply under the style and the paragraph', async () => {
    const { doc, content } = await convertBody(
        styled('Styled', run('Both'), tabs(tab('left', 2880))), {
            styles: '<w:docDefaults><w:pPrDefault><w:pPr>' + tabs(tab('left', 360)) + '</w:pPr></w:pPrDefault></w:docDefaults>' +
                paragraphStyle('Styled', { pPr: tabs(tab('center', 1440)) })
        });

    assert.equal(tabSetOf(doc, content, 'Both'), '18.0:0:0,72.0:1:0,144.0:0:0');
});

test('a clear stop removes the inherited stop at its position, along the chain and on the paragraph', async () => {
    const { doc, content } = await convertBody(
        styled('Child', run('Cleared one'), tabs(tab('clear', 720))) +
        styled('Child', run('Cleared all'), tabs(tab('clear', 720), tab('clear', 2880))) +
        styled('Child', run('Clear elsewhere'), tabs(tab('clear', 1000))), {
            styles: paragraphStyle('Base', { pPr: tabs(tab('left', 720), tab('left', 1440)) }) +
                paragraphStyle('Child', { basedOn: 'Base', pPr: tabs(tab('clear', 1440), tab('right', 2880)) })
        });

    assert.equal(tabSetOf(doc, content, 'Cleared one'), '144.0:2:0');
    assert.equal(tabSetOf(doc, content, 'Cleared all'), null, 'no stops left: no TabSet');
    assert.equal(tabSetOf(doc, content, 'Clear elsewhere'), '36.0:0:0,144.0:2:0');
});

test('stops the paragraph adds merge with the inherited ones in position order', async () => {
    const { doc, content } = await convertBody(
        styled('Styled', run('Merged'), tabs(tab('left', 1440), tab('right', 360), tab('center', 2880))), {
            styles: paragraphStyle('Styled', { pPr: tabs(tab('left', 720), tab('left', 2880)) })
        });

    // The paragraph's center stop at 2880 replaces the style's left stop there
    assert.equal(tabSetOf(doc, content, 'Merged'), '18.0:2:0,36.0:0:0,72.0:0:0,144.0:1:0');
});

test('bar stops and stops with an unreadable or negative position are left out', async () => {
    const { doc, content } = await convertBody(
        p(run('Mixed'), tabs(tab('bar', 720), tab('left', 'abc'), tab('left', -360), tab('left', 1440),
            '<w:tab w:val="left"/>', tab('bogus', 2000))) +
        p(run('Only dropped'), tabs(tab('bar', 720), tab('left', -360), tab('left', 'x'))) +
        p(run('Empty tabs'), tabs()));

    assert.equal(tabSetOf(doc, content, 'Mixed'), '72.0:0:0');
    assert.equal(tabSetOf(doc, content, 'Only dropped'), null);
    assert.equal(tabSetOf(doc, content, 'Empty tabs'), null);
});

test('a position with a unit is converted to points', async () => {
    const { doc, content } = await convertBody(p(run('Inches'), tabs(tab('left', '1in'), tab('left', '2.5cm'))));
    assert.equal(tabSetOf(doc, content, 'Inches'), '70.9:0:0,72.0:0:0');
});

test('paragraphs in table cells carry their tab stops', async () => {
    const { doc, content } = await convertBody(
        '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>' +
        p(run('In cell'), tabs(tab('right', 3600, 'dot'))) +
        '</w:tc></w:tr></w:tbl>');

    const para = [...doc.querySelectorAll('elements cell paragraph')]
        .find(el => [...el.querySelectorAll('content')].some(c => textOf(c, content) === 'In cell'));
    assert.ok(para, 'the cell paragraph is written');
    assert.equal(para.getAttribute('TabSet'), '180.0:2:1');
});

test('a paragraph split by a page break keeps its stops on both parts', async () => {
    const { doc, content } = await convertBody(
        p(run('Before') + '<w:r><w:br w:type="page"/></w:r>' + run('After'), tabs(tab('left', 720))));

    assert.equal(tabSetOf(doc, content, 'Before'), '36.0:0:0');
    assert.equal(tabSetOf(doc, content, 'After'), '36.0:0:0');
});

test('a reused parser does not carry style tab stops into the next document', async () => {
    const { loadConverter, buildDocx } = require('./helpers');
    const { DocxParser } = loadConverter();
    const parser = new DocxParser();

    const first = await parser.parse(await buildDocx(styled('Styled', run('A')),
        { styles: paragraphStyle('Styled', { pPr: tabs(tab('left', 720)) }) }));
    assert.equal(first.elements[0].tabStops.length, 1);

    const second = await parser.parse(await buildDocx(styled('Styled', run('B'))));
    assert.equal(second.elements[0].tabStops, undefined);
});
