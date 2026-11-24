const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

// Track Changes: a document is converted as if every change were accepted, and the changes that
// haven't been accepted are reported (trackedChanges) so the user can check them in Word. A paragraph
// deleted whole and a deleted table row are left out, as Word's Accept All removes them.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const REV = 'w:author="Avukat" w:date="2026-01-01T00:00:00Z"';
const ins = text => `<w:ins w:id="1" ${REV}>${r(text)}</w:ins>`;
const del = text => `<w:del w:id="2" ${REV}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const moveFrom = text => `<w:moveFrom w:id="3" ${REV}>${r(text)}</w:moveFrom>`;
const moveTo = text => `<w:moveTo w:id="4" ${REV}>${r(text)}</w:moveTo>`;
// The paragraph mark's own revision, in w:pPr/w:rPr
const markDeleted = `<w:rPr><w:del w:id="5" ${REV}/></w:rPr>`;
const markInserted = `<w:rPr><w:ins w:id="6" ${REV}/></w:rPr>`;
// A paragraph deleted whole: its mark and its runs
const deletedParagraph = text => p(del(text), markDeleted);
const hidden = text => r(text, '<w:vanish/>');

// A table on a 2-column grid with all borders; each row is an array of cells, each cell its inner
// XML (paragraphs), or { tcPr, content }. trPr, when given, goes in the row's w:trPr.
const tbl = (rows, { tblPr = '' } = {}) => '<w:tbl>' +
    `<w:tblPr>${tblPr}<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
    '<w:bottom w:val="double" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="dotted" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    rows.map(row => {
        const { trPr = '', cells } = Array.isArray(row) ? { cells: row } : row;
        return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}` + cells.map(cell => {
            const { tcPr = '', content } = typeof cell === 'string' ? { content: cell } : cell;
            return `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${content}</w:tc>`;
        }).join('') + '</w:tr>';
    }).join('') + '</w:tbl>';
const rowDeleted = `<w:del w:id="7" ${REV}/>`;
const rowInserted = `<w:ins w:id="8" ${REV}/>`;
const deletedRow = (...texts) => ({ trPr: rowDeleted, cells: texts.map(text => p(del(text))) });

// A floating text box the way Word writes it, DrawingML with a VML fallback
const textBox = inner => '<w:r><mc:AlternateContent>' +
    '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
    '</mc:AlternateContent></w:r>';

const headerPart = inner => ({ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner });
const HEADER_REFERENCE = '<w:headerReference w:type="default" r:id="rIdHeader1"/>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code = 'trackedChanges') => plain(parsed.findings).find(f => f.code === code);
const bodyParagraphs = (doc, content) => [...doc.querySelectorAll('elements > paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));
const rowTexts = (doc, content) => [...doc.querySelectorAll('elements > table > row')]
    .map(row => [...row.querySelectorAll('cell')].map(cell =>
        [...cell.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join('')));

/**
 * Assert that a body converts to exactly the UDF of another body: what Word's Accept All leaves
 * @param {string} withChanges - Body XML with tracked changes
 * @param {string} accepted - The same body with every change accepted by hand
 * @param {Object} [options] - See buildDocx, for both
 */
async function assertConvertsLike(withChanges, accepted, options) {
    const [a, b] = await Promise.all([convertBody(withChanges, options), convertBody(accepted, options)]);
    assert.equal(a.xml, b.xml);
    return a;
}

test('a document without revisions reports no trackedChanges finding', async () => {
    const { parsed } = await convertBody(p(r('Metin')) + tbl([[p(r('A')), p(r('B'))]]) + p(r('Son')));
    assert.equal(finding(parsed), undefined);
});

test('inserted, deleted and moved runs are converted as accepted, and each paragraph is reported', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Davacı ') + ins('kesinlikle ') + r('haklıdır.')) +
        p(r('Faiz ') + del('yüksek ') + r('talep edilir.')) +
        p(r('Değişmeyen paragraf.')) +
        p(r('Önce ') + moveFrom('taşınan ') + r('sonra.')) +
        p(r('Buraya ') + moveTo('taşınan ') + r('metin.')));
    assert.deepEqual(bodyParagraphs(doc, content),
        ['Davacı kesinlikle haklıdır.', 'Faiz talep edilir.', 'Değişmeyen paragraf.', 'Önce sonra.', 'Buraya taşınan metin.']);
    assert.ok(!content.includes('yüksek'));
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 4,
        examples: ['Davacı kesinlikle haklıdır.', 'Faiz talep edilir.', 'Önce sonra.'],
        occurrences: [
            { example: 'Davacı kesinlikle haklıdır.', location: { kind: 'paragraph', paragraph: 1 } },
            { example: 'Faiz talep edilir.', location: { kind: 'paragraph', paragraph: 2 } },
            { example: 'Önce sonra.', location: { kind: 'paragraph', paragraph: 4 } },
            { example: 'Buraya taşınan metin.', location: { kind: 'paragraph', paragraph: 5 } }
        ]
    });
});

test('formatting changes, an inserted paragraph mark and inserted numbering are reported; the output is the new formatting', async () => {
    const rPrChange = `<w:rPrChange w:id="9" ${REV}><w:rPr/></w:rPrChange>`;
    const pPrChange = `<w:pPrChange w:id="10" ${REV}><w:pPr/></w:pPrChange>`;
    const { parsed } = await assertConvertsLike(
        p(r('Kalın yapıldı', `<w:b/>${rPrChange}`)) + p(r('Ortalandı'), `<w:jc w:val="center"/>${pPrChange}`) +
        p(r('Yeni paragraf'), markInserted),
        p(r('Kalın yapıldı', '<w:b/>')) + p(r('Ortalandı'), '<w:jc w:val="center"/>') + p(r('Yeni paragraf')));
    assert.deepEqual(finding(parsed).occurrences.map(o => [o.example, o.location.paragraph]),
        [['Kalın yapıldı', 1], ['Ortalandı', 2], ['Yeni paragraf', 3]]);
    assert.equal(finding(parsed).count, 3);
});

test('a paragraph whose mark and runs are all deleted is left out; the paragraphs around it are unchanged', async () => {
    const { doc, content, parsed } = await assertConvertsLike(
        p(r('Birinci')) + deletedParagraph('Silinen paragraf') + p(r('Üçüncü')),
        p(r('Birinci')) + p(r('Üçüncü')));
    assert.deepEqual(bodyParagraphs(doc, content), ['Birinci', 'Üçüncü']);
    assert.ok(!content.includes('Silinen'));
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 1, examples: [],
        occurrences: [{ example: '', location: { kind: 'paragraph', paragraph: 2 } }]
    });
});

test('an empty paragraph whose mark is deleted is left out too', async () => {
    await assertConvertsLike(p(r('Birinci')) + p('', markDeleted) + p(r('Üçüncü')), p(r('Birinci')) + p(r('Üçüncü')));
});

test('a paragraph with a deleted mark that keeps some text is still written, with that text', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Birinci')) + p(r('Kalan metin') + del(' ve silinen'), markDeleted) + p(r('Üçüncü')));
    assert.deepEqual(bodyParagraphs(doc, content), ['Birinci', 'Kalan metin', 'Üçüncü']);
    assert.deepEqual(finding(parsed).occurrences, [{ example: 'Kalan metin', location: { kind: 'paragraph', paragraph: 2 } }]);
});

test('a paragraph whose mark is deleted but which holds no deleted run and only hidden text is left out', async () => {
    const { content, parsed } = await convertBody(p(r('Birinci')) + p(hidden('gizli'), markDeleted) + p(r('Üçüncü')));
    assert.ok(!content.includes('gizli'));
    assert.equal(finding(parsed).count, 1);
    // The hidden text is still noted, as for any hidden run
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences, [{ example: 'gizli', location: { kind: 'paragraph', paragraph: 2 } }]);
});

test('a moved-away paragraph (its mark in w:moveFrom) is left out where it was and kept where it went', async () => {
    const moveFromMark = `<w:rPr><w:moveFrom w:id="11" ${REV}/></w:rPr>`;
    const moveToMark = `<w:rPr><w:moveTo w:id="12" ${REV}/></w:rPr>`;
    await assertConvertsLike(
        p(moveFrom('Taşınan'), moveFromMark) + p(r('Orta')) + p(moveTo('Taşınan'), moveToMark),
        p(r('Orta')) + p(r('Taşınan')));
});

test('finding locations after a removed paragraph keep Word\'s ¶ numbering', async () => {
    const { parsed } = await convertBody(
        p(r('Birinci')) + deletedParagraph('Silinen') + deletedParagraph('Bir silinen daha') + p(r('Dördüncü') + hidden('gizli')));
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences, [{ example: 'gizli', location: { kind: 'paragraph', paragraph: 4 } }]);
    assert.deepEqual(finding(parsed).occurrences.map(o => o.location.paragraph), [2, 3]);
});

test('a deleted heading names no place after it', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
        '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';
    const heading = (runs, pPr = '') => p(runs, '<w:pStyle w:val="Heading1"/>' + pPr);
    const { parsed } = await convertBody(
        heading(r('AÇIKLAMALAR')) + heading(del('ESKİ BAŞLIK'), markDeleted) + p(r('Metin') + hidden('gizli')), { styles });
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences,
        [{ example: 'gizli', location: { kind: 'paragraph', paragraph: 3, heading: 'AÇIKLAMALAR' } }]);
});

test('a deleted row is left out; the other rows, their cells, merges and borders are unchanged', async () => {
    // The merge from row 1 runs through the deleted row into row 3
    const merged = (tcPr, content) => ({ tcPr, content });
    const { doc, content, parsed } = await assertConvertsLike(
        tbl([
            [merged('<w:vMerge w:val="restart"/>', p(r('Birleşik'))), p(r('A1'))],
            { trPr: rowDeleted, cells: [merged('<w:vMerge/>', p('')), p(del('Silinen satır'))] },
            [merged('<w:vMerge/>', p('')), p(r('A3'))],
            [p(r('B1')), p(r('B2'))]
        ]),
        tbl([
            [merged('<w:vMerge w:val="restart"/>', p(r('Birleşik'))), p(r('A1'))],
            [merged('<w:vMerge/>', p('')), p(r('A3'))],
            [p(r('B1')), p(r('B2'))]
        ]));
    assert.equal(doc.querySelectorAll('elements > table > row').length, 3);
    assert.ok(!content.includes('Silinen'));
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 1, examples: [],
        occurrences: [{ example: '', location: { kind: 'cell', table: 1, row: 2, cell: 1 } }]
    });
});

test('a deleted first or last row gives its borders to the row that takes its place', async () => {
    await assertConvertsLike(
        tbl([deletedRow('X', 'Y'), [p(r('A')), p(r('B'))], [p(r('C')), p(r('D'))], deletedRow('Z', 'W')]),
        tbl([[p(r('A')), p(r('B'))], [p(r('C')), p(r('D'))]]));
});

test('a row whose cells keep text is written even when its text was edited; only a row marked deleted goes', async () => {
    const { doc, content, parsed } = await convertBody(tbl([
        [p(r('Kalan') + del(' silinen')), p(r('Metin'))],
        deletedRow('Satır', 'silindi'),
        [p(r('Son')), p(r('Satır'))]
    ]));
    assert.deepEqual(rowTexts(doc, content), [['Kalan', 'Metin'], ['Son', 'Satır']]);
    // Row 1's edited paragraph, then row 2, located by Word's row numbers
    assert.deepEqual(finding(parsed).occurrences, [
        { example: 'Kalan', location: { kind: 'cell', table: 1, row: 1, cell: 1 } },
        { example: '', location: { kind: 'cell', table: 1, row: 2, cell: 1 } }
    ]);
    assert.equal(finding(parsed).count, 2);
});

test('rows after a deleted row keep Word\'s row numbers in finding locations', async () => {
    const { parsed } = await convertBody(tbl([
        [p(r('A')), p(r('B'))],
        deletedRow('X', 'Y'),
        [p(r('C')), p(r('D') + hidden('gizli'))]
    ]));
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences,
        [{ example: 'gizli', location: { kind: 'cell', table: 1, row: 3, cell: 2 } }]);
});

test('a table whose rows are all deleted is left out, and the next table keeps its number', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Önce')) + tbl([deletedRow('A', 'B'), deletedRow('C', 'D')]) + p(r('Sonra')) +
        tbl([[p(r('E')), p(r('F') + hidden('gizli'))]]));
    assert.equal(doc.querySelectorAll('elements > table').length, 1);
    assert.deepEqual(bodyParagraphs(doc, content), ['Önce', 'Sonra']);
    assert.deepEqual(rowTexts(doc, content), [['E', 'F']]);
    assert.deepEqual(finding(parsed).occurrences, [
        { example: '', location: { kind: 'cell', table: 1, row: 1, cell: 1 } },
        { example: '', location: { kind: 'cell', table: 1, row: 2, cell: 1 } }
    ]);
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences,
        [{ example: 'gizli', location: { kind: 'cell', table: 2, row: 1, cell: 2 } }]);
});

test('an inserted row is kept and reported at its first cell, besides its inserted text', async () => {
    const { doc, content, parsed } = await convertBody(tbl([
        [p(r('A')), p(r('B'))],
        { trPr: rowInserted, cells: [p(ins('Yeni')), p(ins('satır'))] }
    ]));
    assert.deepEqual(rowTexts(doc, content), [['A', 'B'], ['Yeni', 'satır']]);
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 3, examples: ['Yeni', 'satır'],
        occurrences: [
            { example: 'Yeni', location: { kind: 'cell', table: 1, row: 2, cell: 1 } },
            { example: 'satır', location: { kind: 'cell', table: 1, row: 2, cell: 2 } }
        ]
    });
});

test('cell changes (inserted, deleted and merged cells, cell formatting changes) are reported at their cell', async () => {
    const { parsed } = await convertBody(tbl([
        [{ tcPr: `<w:cellIns w:id="13" ${REV}/>`, content: p(r('Eklenen hücre')) },
            { tcPr: `<w:tcPrChange w:id="14" ${REV}><w:tcPr/></w:tcPrChange>`, content: p(r('Biçim')) }],
        [{ tcPr: `<w:cellDel w:id="15" ${REV}/>`, content: p(r('Silinen hücre')) },
            { tcPr: `<w:cellMerge w:id="16" ${REV} w:vMerge="cont"/>`, content: p('') }]
    ]));
    assert.deepEqual(finding(parsed).occurrences.map(o => [o.example, o.location.row, o.location.cell]),
        [['Eklenen hücre', 1, 1], ['Biçim', 1, 2], ['Silinen hücre', 2, 1], ['', 2, 2]]);
});

test('a row formatting change is reported; the row is kept', async () => {
    const { doc, content, parsed } = await convertBody(tbl([
        { trPr: `<w:trHeight w:val="400"/><w:trPrChange w:id="17" ${REV}><w:trPr/></w:trPrChange>`, cells: [p(r('A')), p(r('B'))] }
    ]));
    assert.deepEqual(rowTexts(doc, content), [['A', 'B']]);
    assert.deepEqual(finding(parsed).occurrences, [{ example: 'A', location: { kind: 'cell', table: 1, row: 1, cell: 1 } }]);
});

test('a deleted paragraph in a cell is left out; a cell left with no paragraph is written empty', async () => {
    await assertConvertsLike(
        tbl([[p(r('Kalan')) + deletedParagraph('Silinen'), deletedParagraph('Tek paragraf')]]),
        tbl([[p(r('Kalan')), p('')]]));
});

test('revisions in a cell that continues a vertical merge, written empty, are not reported', async () => {
    const { parsed } = await convertBody(tbl([
        [{ tcPr: '<w:vMerge w:val="restart"/>', content: p(r('Birleşik')) }, p(r('A'))],
        [{ tcPr: '<w:vMerge/>', content: p(ins('görünmez')) }, p(r('B'))]
    ]));
    assert.equal(finding(parsed), undefined);
});

test('a deleted row in a nested table is left out, and reported at the outer cell', async () => {
    const nested = tbl([[p(r('İç A')), p(r('İç B'))], deletedRow('İç silinen', 'satır')]);
    const { parsed } = await assertConvertsLike(
        tbl([[p(r('Dış')), nested + p('')]]),
        tbl([[p(r('Dış')), tbl([[p(r('İç A')), p(r('İç B'))]]) + p('')]]));
    assert.deepEqual(finding(parsed).occurrences, [{ example: '', location: { kind: 'cell', table: 1, row: 1, cell: 2 } }]);
});

test('a nested table whose rows are all deleted is left out of its cell', async () => {
    await assertConvertsLike(
        tbl([[p(r('Dış')), tbl([deletedRow('A', 'B')]) + p(r('Sonra'))]]),
        tbl([[p(r('Dış')), p(r('Sonra'))]]));
});

test('a deleted paragraph in a text box is left out, and reported at the anchoring paragraph', async () => {
    const { content, parsed } = await convertBody(
        p(r('Önce')) + p(r('Çapa') + textBox(p(r('Kutu metni')) + deletedParagraph('Kutudan silinen'))));
    assert.ok(content.includes('Kutu metni') && !content.includes('Kutudan silinen'));
    assert.deepEqual(finding(parsed).occurrences, [{ example: '', location: { kind: 'paragraph', paragraph: 2 } }]);
    // The same text box without the deleted paragraph
    await assertConvertsLike(
        p(r('Önce')) + p(r('Çapa') + textBox(p(r('Kutu metni')) + deletedParagraph('Kutudan silinen'))),
        p(r('Önce')) + p(r('Çapa') + textBox(p(r('Kutu metni')))));
});

test('a paragraph whose mark and runs are deleted but which anchors a text box is left out; the text box is kept', async () => {
    await assertConvertsLike(
        p(r('Önce')) + p(del('Silinen') + textBox(p(r('Kutu'))), markDeleted) + p(r('Sonra')),
        p(r('Önce')) + p(textBox(p(r('Kutu')))) + p(r('Sonra')));
});

test('in the converted header, a deleted paragraph is left out and changes are reported at the header', async () => {
    const options = inner => ({ parts: [headerPart(inner)], sectPr: HEADER_REFERENCE });
    const withChanges = p(r('Hukuk Bürosu') + ins(' Ankara')) + deletedParagraph('Eski adres');
    const { doc, content, parsed } = await convertBody(p(r('Gövde')), options(withChanges));
    const header = doc.querySelector('header');
    assert.equal(header.querySelectorAll('paragraph').length, 1);
    assert.equal([...header.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''), 'Hukuk Bürosu Ankara');
    assert.ok(!content.includes('Eski adres'));
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 2, examples: ['Hukuk Bürosu Ankara'],
        occurrences: [
            { example: 'Hukuk Bürosu Ankara', location: { kind: 'header' } },
            { example: '', location: { kind: 'header' } }
        ]
    });
    const accepted = await convertBody(p(r('Gövde')), options(p(r('Hukuk Bürosu') + r(' Ankara'))));
    assert.equal(content, accepted.content);
});

test('a converted header emptied by unaccepted deletions is reported, and nothing of it is written', async () => {
    const { doc, parsed } = await convertBody(p(r('Gövde')),
        { parts: [headerPart(deletedParagraph('Eski büro adı') + p(del('Eski adres')))], sectPr: HEADER_REFERENCE });
    assert.equal(doc.querySelector('header'), null);
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 2, examples: [],
        occurrences: [{ example: '', location: { kind: 'header' } }]
    });
});

test('a change to the final section\'s properties is reported for the whole document', async () => {
    const sectPr = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>' +
        `<w:sectPrChange w:id="18" ${REV}><w:sectPr><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:sectPrChange>`;
    const { parsed } = await convertBody(p(r('Metin')), { sectPr });
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 1, examples: [],
        occurrences: [{ example: '', location: { kind: 'document' } }]
    });
    // The margins written are the current ones
    assert.equal(parsed.pageSetup.leftMargin, 70.9);
});

test('in a footnote, deleted and moved-away text and a deleted paragraph are left out, and reported at the reference', async () => {
    const note = p('<w:r><w:footnoteRef/></w:r>' + r(' Dipnot ') + ins('eklenen ') + del('silinen ') + moveFrom('taşınan ') + r('metni.')) +
        deletedParagraph('Silinen dipnot paragrafı') + p(r('İkinci paragraf.'));
    const { content, parsed } = await convertBody(p(r('Metin') + '<w:r><w:footnoteReference w:id="1"/></w:r>'), { footnotes: { 1: note } });
    assert.ok(content.includes('Dipnot eklenen metni.'));
    assert.ok(content.includes('İkinci paragraf.'));
    for (const gone of ['silinen ', 'taşınan', 'Silinen dipnot']) assert.ok(!content.includes(gone), gone);
    assert.deepEqual(finding(parsed, 'notesMoved').examples, ['Dipnot eklenen metni. İkinci paragraf.']);
    const location = { kind: 'paragraph', paragraph: 1, note: { kind: 'footnote' } };
    assert.deepEqual(finding(parsed), {
        code: 'trackedChanges', count: 2, examples: ['Dipnot eklenen metni.'],
        occurrences: [{ example: 'Dipnot eklenen metni.', location }, { example: '', location }]
    });
    // Exactly the UDF of the note with every change accepted
    const accepted = p('<w:r><w:footnoteRef/></w:r>' + r(' Dipnot ') + r('eklenen ') + r('metni.')) + p(r('İkinci paragraf.'));
    const expected = await convertBody(p(r('Metin') + '<w:r><w:footnoteReference w:id="1"/></w:r>'), { footnotes: { 1: accepted } });
    assert.equal(content, expected.content);
});

for (const kind of ['footnote', 'endnote']) {
    test(`a deleted row in a ${kind} is omitted from the export and note text`, async () => {
        const body = p(r('Metin') + `<w:r><w:${kind}Reference w:id="1"/></w:r>`);
        const kept = [p(r('Güncel tutar')), p(r('200 TL'))];
        const removed = { trPr: rowDeleted, cells: [p(r('Eski tutar')), p(r('100 TL'))] };
        const note = rows => ({ [`${kind}s`]: { 1: p(`<w:r><w:${kind}Ref/></w:r>`) + tbl(rows) } });
        const { xml, parsed } = await convertBody(body, note([kept, removed]));
        assert.equal(xml, (await convertBody(body, note([kept]))).xml);
        assert.equal(parsed[`${kind}s`][1], 'Güncel tutar 200 TL');
        assert.deepEqual(finding(parsed), {
            code: 'trackedChanges', count: 1, examples: [],
            occurrences: [{ example: '', location: { kind: 'paragraph', paragraph: 1, note: { kind } } }]
        });
    });
}

test('the note only reports: the .udf is the same with or without it', async () => {
    const { loadConverter, buildDocx } = require('./helpers');
    const JSZip = require('jszip');
    const { DocxParser, UdfGenerator } = loadConverter();
    const parsed = await new DocxParser().parse(await buildDocx(p(r('A') + ins('B')) + deletedParagraph('C')));
    assert.equal(parsed.findings.length, 1);
    const udfXml = async document => (await JSZip.loadAsync(
        await (await new UdfGenerator().generate(document)).arrayBuffer())).file('content.xml').async('string');
    assert.equal(await udfXml(parsed), await udfXml({ ...parsed, findings: [] }));
});
