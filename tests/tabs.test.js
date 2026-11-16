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

test('a clear or a bar stop one twip from an inherited stop leaves that stop in place', async () => {
    // 1441 and 1442 twips are 72.05 and 72.1 points: different stops, though both are written as 72.1
    const { doc, content } = await convertBody(
        styled('Styled', run('Clear beside'), tabs(tab('clear', 1441))) +
        styled('Styled', run('Clear on it'), tabs(tab('clear', 1442))) +
        styled('Styled', run('Stop on it'), tabs(tab('right', 1442, 'dot'))) +
        styled('Styled', run('Bar beside'), tabs(tab('bar', 1441))) +
        styled('Styled', run('Bar on it'), tabs(tab('bar', 1442))), {
            styles: paragraphStyle('Styled', { pPr: tabs(tab('left', 1442)) })
        });

    assert.equal(tabSetOf(doc, content, 'Clear beside'), '72.1:0:0');
    assert.equal(tabSetOf(doc, content, 'Clear on it'), null);
    assert.equal(tabSetOf(doc, content, 'Stop on it'), '72.1:2:1');
    assert.equal(tabSetOf(doc, content, 'Bar beside'), '72.1:0:0');
    assert.equal(tabSetOf(doc, content, 'Bar on it'), null);
});

test('two stops left at the same written position become one: the first in position order', async () => {
    // 1441 and 1442 twips are both written as 72.1, and 1439 and 1440 both as 72.0. TabSet has
    // no meaning for two stops at one position, so only the stop a tab reaches first is written.
    const { doc, content, parsed } = await convertBody(
        styled('Upper', run('Paragraph stop first'), tabs(tab('right', 1441, 'dot'))) +
        styled('Lower', run('Style stop first'), tabs(tab('right', 1442, 'dot'))) +
        p(run('Listed backwards'), tabs(tab('center', 1440), tab('decimal', 1439, 'hyphen'), tab('left', 720))) +
        p(run('Two pairs'), tabs(tab('left', 1443), tab('center', 1442), tab('right', 1441), tab('left', 1444))) +
        p(run('A twip apart, two tenths'), tabs(tab('left', 1440), tab('right', 1441))), {
            styles: paragraphStyle('Upper', { pPr: tabs(tab('left', 1442)) }) +
                paragraphStyle('Lower', { pPr: tabs(tab('left', 1441)) })
        });

    assert.equal(tabSetOf(doc, content, 'Paragraph stop first'), '72.1:2:1');
    assert.equal(tabSetOf(doc, content, 'Style stop first'), '72.1:0:0');
    assert.equal(tabSetOf(doc, content, 'Listed backwards'), '36.0:0:0,72.0:3:2');
    // 1441, 1442, 1443 and 1444 twips are 72.05, 72.1, 72.15 and 72.2 points
    assert.equal(tabSetOf(doc, content, 'Two pairs'), '72.1:2:0,72.2:0:0');
    assert.equal(tabSetOf(doc, content, 'A twip apart, two tenths'), '72.0:0:0,72.1:2:0');
    assert.deepEqual(JSON.parse(JSON.stringify(parsed.elements[0].tabStops)), [{ position: 72.1, alignment: 2, leader: 1 }]);

    // No paragraph is written with a repeated position
    for (const paragraph of doc.querySelectorAll('elements paragraph[TabSet]')) {
        const positions = paragraph.getAttribute('TabSet').split(',').map(stop => stop.split(':')[0]);
        assert.equal(new Set(positions).size, positions.length, paragraph.getAttribute('TabSet'));
    }
});

test('a clear, a stop or a bar stop removes the stop at the same position written in another unit', async () => {
    // The same lengths, which convert to different floating-point numbers of points:
    // 432 twips is 21.6 but 0.3in is 21.599999999999998, and 7200 twips is 360 but 127mm is 360.00000000000006
    const equal = [[432, '0.3in'], [7200, '127mm'], [9360, '16.51cm'], [264, '1.1pc'], ['0.6in', '43.2pt'], ['0.3in', 432]];
    const body = equal.map(([inherited, own], i) =>
        styled(`S${i}`, run(`Clear ${i}`), tabs(tab('clear', own))) +
        styled(`S${i}`, run(`Stop ${i}`), tabs(tab('right', own, 'dot'))) +
        styled(`S${i}`, run(`Bar ${i}`), tabs(tab('bar', own))) +
        styled(`S${i}`, run(`Inherited ${i}`))).join('');
    const { doc, content } = await convertBody(body, {
        styles: equal.map(([inherited], i) => paragraphStyle(`S${i}`, { pPr: tabs(tab('left', inherited)) })).join('')
    });

    equal.forEach(([inherited, own], i) => {
        const written = tabSetOf(doc, content, `Inherited ${i}`);
        assert.match(written, /^\d+\.\d:0:0$/, `${inherited}`);
        assert.equal(tabSetOf(doc, content, `Clear ${i}`), null, `clear at ${own} over a stop at ${inherited}`);
        assert.equal(tabSetOf(doc, content, `Stop ${i}`), written.replace(':0:0', ':2:1'), `stop at ${own} over a stop at ${inherited}`);
        assert.equal(tabSetOf(doc, content, `Bar ${i}`), null, `bar at ${own} over a stop at ${inherited}`);
    });
});

test('stops the paragraph adds merge with the inherited ones in position order', async () => {
    const { doc, content } = await convertBody(
        styled('Styled', run('Merged'), tabs(tab('left', 1440), tab('right', 360), tab('center', 2880))), {
            styles: paragraphStyle('Styled', { pPr: tabs(tab('left', 720), tab('left', 2880)) })
        });

    // The paragraph's center stop at 2880 replaces the style's left stop there
    assert.equal(tabSetOf(doc, content, 'Merged'), '18.0:2:0,36.0:0:0,72.0:0:0,144.0:1:0');
});

test('bar stops replace inherited text stops from defaults and styles without restoring them', async () => {
    const { doc, content, parsed } = await convertBody(
        p(run('Default replaced'), tabs(tab('bar', 720))) +
        styled('Base', run('Style replaced'), tabs(tab('bar', 1440))) +
        styled('BarDefault', run('Style bar')) +
        styled('BarChild', run('Child bar')), {
            styles: '<w:docDefaults><w:pPrDefault><w:pPr>' + tabs(tab('right', 720, 'dot')) +
                '</w:pPr></w:pPrDefault></w:docDefaults>' +
                paragraphStyle('Base', { pPr: tabs(tab('center', 1440)) }) +
                paragraphStyle('BarDefault', { pPr: tabs(tab('bar', 720)) }) +
                paragraphStyle('BarChild', { basedOn: 'Base', pPr: tabs(tab('bar', 1440)) })
        });

    assert.equal(tabSetOf(doc, content, 'Default replaced'), null);
    assert.equal(tabSetOf(doc, content, 'Style replaced'), '36.0:2:1');
    assert.equal(tabSetOf(doc, content, 'Style bar'), null);
    assert.equal(tabSetOf(doc, content, 'Child bar'), '36.0:2:1');
    assert.deepEqual(JSON.parse(JSON.stringify(parsed.findings)), [
        { code: 'tabStops', count: 4, examples: ['Default replaced', 'Style replaced', 'Style bar'] }
    ]);
});

test('a later normal stop replaces an inherited bar, while a clear removes it', async () => {
    const { doc, content, parsed } = await convertBody(
        styled('Bar', run('Normal override'), tabs(tab('center', 720, 'hyphen'))) +
        styled('Bar', run('Clear override'), tabs(tab('clear', 720))) +
        styled('Normal', run('Child normal')), {
            styles: '<w:docDefaults><w:pPrDefault><w:pPr>' + tabs(tab('right', 720, 'dot')) +
                '</w:pPr></w:pPrDefault></w:docDefaults>' +
                paragraphStyle('Bar', { pPr: tabs(tab('bar', 720)) }) +
                paragraphStyle('Normal', { basedOn: 'Bar', pPr: tabs(tab('left', 720)) })
        });

    assert.equal(tabSetOf(doc, content, 'Normal override'), '36.0:1:2');
    assert.equal(tabSetOf(doc, content, 'Clear override'), null);
    assert.equal(tabSetOf(doc, content, 'Child normal'), '36.0:0:0');
    assert.deepEqual(JSON.parse(JSON.stringify(parsed.findings)), []);
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
