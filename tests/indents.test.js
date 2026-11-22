const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const run = text => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const TAB = '<w:r><w:tab/></w:r>';
const ind = attrs => `<w:ind ${Object.entries(attrs).map(([name, value]) => `w:${name}="${value}"`).join(' ')}/>`;
const tabs = (...stops) => `<w:tabs>${stops.join('')}</w:tabs>`;
const tab = (val, pos) => `<w:tab w:val="${val}" w:pos="${pos}"/>`;
const styled = (styleId, runs, pPr = '') => p(runs, `<w:pStyle w:val="${styleId}"/>${pPr}`);
const paragraphStyle = (id, pPr) => `<w:style w:type="paragraph" w:styleId="${id}"><w:pPr>${pPr}</w:pPr></w:style>`;
const tbl = inner => '<w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
    `<w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;
const numPr = numId => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;
const numbering = '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="none"/><w:lvlText w:val=""/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>' +
    '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>';

// A party line: DAVACI<tab>: Ahmet
const partyRuns = label => run(label) + TAB + run(': Ahmet');

/** The UDF paragraph that holds this text in exactly one content element */
function paragraphOf(doc, content, text) {
    const found = [...doc.querySelectorAll('elements content[startOffset]')].filter(el => textOf(el, content) === text);
    assert.equal(found.length, 1, `${JSON.stringify(text)} is held by ${found.length} elements`);
    return found[0].closest('paragraph');
}

/** LeftIndent, FirstLineIndent and TabSet of the paragraph holding this text (null when absent) */
function indentsOf(doc, content, text) {
    const para = paragraphOf(doc, content, text);
    return {
        left: para.getAttribute('LeftIndent'),
        firstLine: para.getAttribute('FirstLineIndent'),
        tabSet: para.getAttribute('TabSet')
    };
}

test('a hanging indent is written as a negative FirstLineIndent with the implicit stop at the left indent', async () => {
    const { doc, content, parsed } = await convertBody(p(partyRuns('DAVACI'), ind({ left: 1440, hanging: 1440 })));

    assert.deepEqual(indentsOf(doc, content, 'DAVACI'), { left: '72.0', firstLine: '-72.0', tabSet: '72.0:0:0' });
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('a smaller hanging indent keeps the first line inside the left indent', async () => {
    const { doc, content } = await convertBody(p(partyRuns('VEKİLİ'), ind({ left: 2160, hanging: 720 })));

    assert.deepEqual(indentsOf(doc, content, 'VEKİLİ'), { left: '108.0', firstLine: '-36.0', tabSet: '108.0:0:0' });
});

test('a hanging indent larger than the left indent is clamped so the first line starts at the margin', async () => {
    const { doc, content, parsed } = await convertBody(
        p(run('Clamped'), ind({ left: 360, hanging: 720 })) +
        p(run('No left'), ind({ hanging: 720 })));

    assert.deepEqual(indentsOf(doc, content, 'Clamped'), { left: '18.0', firstLine: '-18.0', tabSet: '18.0:0:0' });
    // Nothing left to hang: no first-line indent and no implicit stop
    assert.deepEqual(indentsOf(doc, content, 'No left'), { left: '0.0', firstLine: null, tabSet: null });
    assert.ok(Object.is(parsed.elements[1].firstLineIndent, 0), 'no negative zero');
});

test('w:hanging wins over w:firstLine on the same w:ind', async () => {
    const { doc, content } = await convertBody(
        p(run('Both'), ind({ left: 1440, firstLine: 360, hanging: 1440 })) +
        p(run('Both, other order'), ind({ left: 1440, hanging: 720, firstLine: 360 })));

    assert.equal(indentsOf(doc, content, 'Both').firstLine, '-72.0');
    assert.equal(indentsOf(doc, content, 'Both, other order').firstLine, '-36.0');
});

test('a hanging indent from the paragraph style chain is written the same way, and the nearer layer wins', async () => {
    const { doc, content } = await convertBody(
        styled('Party', run('From style')) +
        styled('Party', run('Own first line'), ind({ firstLine: 360 })) +
        styled('Indented', run('Own hanging'), ind({ hanging: 720 })) +
        styled('Indented', run('Style first line')), {
            styles: paragraphStyle('Party', ind({ left: 1440, hanging: 1440 })) +
                paragraphStyle('Indented', ind({ left: 1440, firstLine: 360 }))
        });

    assert.deepEqual(indentsOf(doc, content, 'From style'), { left: '72.0', firstLine: '-72.0', tabSet: '72.0:0:0' });
    assert.deepEqual(indentsOf(doc, content, 'Own first line'), { left: '72.0', firstLine: '18.0', tabSet: null });
    assert.deepEqual(indentsOf(doc, content, 'Own hanging'), { left: '72.0', firstLine: '-36.0', tabSet: '72.0:0:0' });
    assert.deepEqual(indentsOf(doc, content, 'Style first line'), { left: '72.0', firstLine: '18.0', tabSet: null });
});

test('a hanging indent from w:docDefaults applies too', async () => {
    const { doc, content } = await convertBody(p(run('Default'), ind({ left: 1440 })), {
        styles: '<w:docDefaults><w:pPrDefault><w:pPr>' + ind({ hanging: 720 }) + '</w:pPr></w:pPrDefault></w:docDefaults>'
    });

    assert.deepEqual(indentsOf(doc, content, 'Default'), { left: '72.0', firstLine: '-36.0', tabSet: '72.0:0:0' });
});

test('an explicit stop between the first line and the left indent replaces the implicit stop', async () => {
    const { doc, content } = await convertBody(
        p(run('Own stop'), tabs(tab('left', 1134)) + ind({ left: 1440, hanging: 1440 })) +
        p(run('Stop past'), tabs(tab('left', 2880)) + ind({ left: 1440, hanging: 1440 })) +
        p(run('Stop at first line'), tabs(tab('left', 720)) + ind({ left: 1440, hanging: 720 })));

    // 1134 twips (56.7 pt) lies between 0 and 72: only the explicit stop
    assert.equal(indentsOf(doc, content, 'Own stop').tabSet, '56.7:0:0');
    // A stop past the left indent doesn't stand in for the implicit one
    assert.equal(indentsOf(doc, content, 'Stop past').tabSet, '72.0:0:0,144.0:0:0');
    // A stop where the first line starts isn't past it
    assert.equal(indentsOf(doc, content, 'Stop at first line').tabSet, '36.0:0:0,72.0:0:0');
});

test('an explicit stop at the left indent is not written twice', async () => {
    const { doc, content } = await convertBody(
        p(run('Same'), tabs(tab('right', 1440)) + ind({ left: 1440, hanging: 1440 })) +
        // The left indent rounds to 113 pt, the stop is at 113.4 pt: the same place
        p(run('Rounded'), tabs(tab('left', 2268)) + ind({ left: 2268, hanging: 2268 })));

    assert.equal(indentsOf(doc, content, 'Same').tabSet, '72.0:2:0');
    assert.deepEqual(indentsOf(doc, content, 'Rounded'), { left: '113.0', firstLine: '-113.0', tabSet: '113.4:0:0' });
});

test('numbered and bulleted paragraphs with a hanging indent are written as before', async () => {
    const { doc, content } = await convertBody(
        p(run('Numbered'), numPr(1) + ind({ left: 720, hanging: 360 })) +
        p(run('Bulleted'), numPr(2) + ind({ left: 720, hanging: 360 })) +
        p(run('No definition'), numPr(7) + ind({ left: 720, hanging: 360 })) +
        p(run('Numbered with stop'), numPr(1) + tabs(tab('left', 2880)) + ind({ left: 720, hanging: 360 })), { numbering });

    for (const text of ['Numbered', 'Bulleted', 'No definition']) {
        const para = paragraphOf(doc, content, text);
        assert.ok(para.getAttribute('ListLevel') !== null, `${text} is a list item`);
        assert.equal(para.getAttribute('LeftIndent'), '36.0', text);
        assert.equal(para.getAttribute('FirstLineIndent'), null, text);
        assert.equal(para.getAttribute('TabSet'), null, text);
    }
    assert.equal(indentsOf(doc, content, 'Numbered with stop').tabSet, '144.0:0:0');
});

test('a numbering level without a number (w:numFmt none) hangs like a body paragraph', async () => {
    const { doc, content } = await convertBody(p(run('Unnumbered'), numPr(3) + ind({ left: 720, hanging: 360 })), { numbering });

    const para = paragraphOf(doc, content, 'Unnumbered');
    assert.equal(para.getAttribute('ListLevel'), null);
    assert.equal(para.getAttribute('FirstLineIndent'), '-18.0');
    assert.equal(para.getAttribute('TabSet'), '36.0:0:0');
});

test('paragraphs in table cells get the same hanging indent and implicit stop', async () => {
    const { doc, content } = await convertBody(tbl(p(partyRuns('DAVALI'), ind({ left: 1440, hanging: 1440 }))));

    const para = paragraphOf(doc, content, 'DAVALI');
    assert.ok(para.closest('cell'), 'the paragraph is in a cell');
    assert.deepEqual(indentsOf(doc, content, 'DAVALI'), { left: '72.0', firstLine: '-72.0', tabSet: '72.0:0:0' });
});

test('a paragraph split by a page break hangs only its first fragment', async () => {
    const { doc, content } = await convertBody(p(
        '<w:r><w:t>First</w:t><w:br w:type="page"/><w:t>Second</w:t></w:r>',
        ind({ left: 1440, hanging: 1440 })));

    assert.equal(indentsOf(doc, content, 'First').firstLine, '-72.0');
    assert.equal(indentsOf(doc, content, 'Second').firstLine, null);
    assert.equal(indentsOf(doc, content, 'Second').left, '72.0');
});

test('paragraphs without a hanging indent are written as before', async () => {
    const { doc, content } = await convertBody(
        p(run('First line'), ind({ left: 720, firstLine: 720 })) +
        p(run('Left only'), ind({ left: 720 })) +
        p(run('Zero hanging'), ind({ left: 720, hanging: 0 })));

    assert.deepEqual(indentsOf(doc, content, 'First line'), { left: '36.0', firstLine: '36.0', tabSet: null });
    assert.deepEqual(indentsOf(doc, content, 'Left only'), { left: '36.0', firstLine: null, tabSet: null });
    assert.deepEqual(indentsOf(doc, content, 'Zero hanging'), { left: '36.0', firstLine: null, tabSet: null });
});
