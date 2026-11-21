const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraphStyle = (id, rPr) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;
const characterStyle = (id, rPr) => `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const header = inner => ({ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner });
const headerRef = '<w:headerReference w:type="default" r:id="rIdHeader1"/>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));

/** Text of each body paragraph (header, footer and table paragraphs excluded), from the CDATA pool */
const bodyParagraphs = (doc, content) => [...doc.querySelectorAll('elements > paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));

/** Text of every paragraph in the document's table cells */
const cellParagraphs = (doc, content) => [...doc.querySelectorAll('cell paragraph')]
    .map(el => [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));

const findingsOf = parsed => findingSummaries(parsed.findings);

// What the generator writes for a paragraph with no text
const EMPTY_PARAGRAPH = '\u200B';

test('a run with direct w:vanish writes no text, and the visible text around it keeps consistent offsets', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Önce ') + r('gizli taslak notu', '<w:vanish/>') + r('sonra', '<w:b/>')) + p(r('İkinci paragraf')));

    assert.equal(content.includes('gizli'), false);
    assert.deepEqual(bodyParagraphs(doc, content), ['Önce sonra', 'İkinci paragraf']);
    assert.deepEqual(findingsOf(parsed), [{ code: 'hiddenText', count: 1, examples: ['gizli taslak notu'] }]);
    // The visible runs keep their own formatting; no run carries the internal properties
    const runs = plain(parsed.elements[0].runs);
    assert.deepEqual(runs.map(run => [run.text, run.bold]), [['Önce ', false], ['sonra', true]]);
    for (const run of runs) {
        for (const key of ['hidden', 'caps', 'smallCaps', 'lang']) assert.equal(key in run, false, key);
    }
});

test('tabs, breaks and note references in a hidden run are left out too', async () => {
    const hiddenRun = '<w:r><w:rPr><w:vanish/></w:rPr><w:tab/><w:t>Not</w:t><w:br/><w:footnoteReference w:id="1"/></w:r>';
    const { doc, content, parsed } = await convertBody(p(r('Metin') + hiddenRun), {
        footnotes: { 1: p(r('Dipnot metni')) }
    });

    assert.deepEqual(plain(parsed.elements[0].runs.map(run => run.type)), ['text']);
    assert.equal(content.includes('Dipnot metni'), false);
    assert.deepEqual(bodyParagraphs(doc, content), ['Metin']);
    assert.deepEqual(findingsOf(parsed), [{ code: 'hiddenText', count: 1, examples: ['Not'] }]);
});

test('hidden text is counted once per paragraph, with the start of the paragraph\'s hidden text as the example', async () => {
    const long = 'Bu uzun gizli açıklama müvekkil ile yapılan görüşmenin notlarını içerir';
    const { parsed } = await convertBody(
        p(r('Bir', '<w:vanish/>') + r(' görünen ') + r('iki', '<w:vanish/>')) +
        p(r(long, '<w:vanish/>')) +
        p(r('Gizlisiz')));

    const [finding] = findingsOf(parsed);
    assert.equal(finding.code, 'hiddenText');
    assert.equal(finding.count, 2);
    assert.equal(finding.examples[0], 'Biriki');
    assert.ok(finding.examples[1].endsWith('…') && long.startsWith(finding.examples[1].slice(0, -1)));
});

test('a paragraph whose runs are all hidden is still written as a paragraph', async () => {
    const { doc, content, parsed } = await convertBody(p(r('Birinci')) + p(r('Tamamen gizli', '<w:vanish/>')) + p(r('Üçüncü')));

    assert.deepEqual(bodyParagraphs(doc, content), ['Birinci', EMPTY_PARAGRAPH, 'Üçüncü']);
    assert.equal(parsed.elements.length, 3);
    assert.deepEqual(findingsOf(parsed), [{ code: 'hiddenText', count: 1, examples: ['Tamamen gizli'] }]);
});

test('an empty hidden run reports nothing', async () => {
    const { parsed } = await convertBody(p(r('Metin') + '<w:r><w:rPr><w:vanish/></w:rPr></w:r>'));
    assert.deepEqual(findingsOf(parsed), []);
});

test('a document without hidden text reports no hiddenText finding', async () => {
    const { parsed } = await convertBody(p(r('Görünür', '<w:vanish w:val="0"/>') + r(' metin', '<w:caps/>')));
    assert.deepEqual(findingsOf(parsed), []);
});

test('hidden text set by a character style or a paragraph style is left out', async () => {
    const styles = characterStyle('Gizli', '<w:vanish/>') + paragraphStyle('GizliParagraf', '<w:vanish/>');
    const { doc, content, parsed } = await convertBody(
        p(r('Görünür ') + r('stil ile gizli', '<w:rStyle w:val="Gizli"/>')) +
        p(r('Paragraf stili ile gizli'), '<w:pStyle w:val="GizliParagraf"/>') +
        p(r('Son')), { styles });

    assert.deepEqual(bodyParagraphs(doc, content), ['Görünür ', EMPTY_PARAGRAPH, 'Son']);
    assert.deepEqual(findingsOf(parsed), [
        { code: 'hiddenText', count: 2, examples: ['stil ile gizli', 'Paragraf stili ile gizli'] }
    ]);
});

test('w:vanish w:val="0" on the run overrides a hidden style and keeps the text', async () => {
    const styles = characterStyle('Gizli', '<w:vanish/>') + paragraphStyle('GizliParagraf', '<w:vanish/>');
    const { doc, content, parsed } = await convertBody(
        p(r('Karakter stili açılmış', '<w:rStyle w:val="Gizli"/><w:vanish w:val="0"/>')) +
        p(r('Paragraf stili açılmış', '<w:vanish w:val="false"/>'), '<w:pStyle w:val="GizliParagraf"/>'), { styles });

    assert.deepEqual(bodyParagraphs(doc, content), ['Karakter stili açılmış', 'Paragraf stili açılmış']);
    assert.deepEqual(findingsOf(parsed), []);
});

test('hidden set by both the paragraph and the character style toggles off, like bold', async () => {
    const styles = characterStyle('Gizli', '<w:vanish/>') + paragraphStyle('GizliParagraf', '<w:vanish/>');
    const { doc, content } = await convertBody(
        p(r('İki stil', '<w:rStyle w:val="Gizli"/>'), '<w:pStyle w:val="GizliParagraf"/>'), { styles });
    assert.deepEqual(bodyParagraphs(doc, content), ['İki stil']);
});

test('hidden text is left out of table cells and the converted header', async () => {
    const { doc, content, parsed } = await convertBody(
        tbl([p(r('Hücre ') + r('gizli hücre', '<w:vanish/>')), p(r('Diğer hücre'))]) + p(r('Gövde')), {
            sectPr: headerRef,
            parts: [header(p(r('Büro adı') + r(' gizli başlık', '<w:vanish/>')))]
        });

    assert.equal(content.includes('gizli'), false);
    assert.deepEqual(cellParagraphs(doc, content), ['Hücre ', 'Diğer hücre']);
    const headerEl = doc.querySelector('elements > header');
    assert.ok(headerEl);
    assert.equal([...headerEl.querySelectorAll('[startOffset]')].map(el => textOf(el, content)).join(''), 'Büro adı');
    assert.deepEqual(findingsOf(parsed), [
        { code: 'hiddenText', count: 2, examples: ['gizli başlık', 'gizli hücre'] }
    ]);
});

test('a header whose only text is hidden writes no header', async () => {
    const { doc, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: headerRef,
        parts: [header(p(r('Gizli başlık', '<w:vanish/>')))]
    });
    assert.equal(doc.querySelector('elements > header'), null);
    assert.equal(parsed.header, null);
    assert.deepEqual(findingsOf(parsed), [{ code: 'hiddenText', count: 1, examples: ['Gizli başlık'] }]);
});

test('All Caps text is written upper-case, the Turkish way when the language is Turkish or unset', async () => {
    const { doc, content } = await convertBody(
        p(r("İstanbul 3. Asliye Hukuk Mahkemesi'ne", '<w:caps/>')) +
        p(r('ışık iğne', '<w:caps/><w:lang w:val="tr-TR"/>')) +
        p(r('indigo ink', '<w:caps/><w:lang w:val="en-US"/>')) +
        p(r('düz metin')));

    assert.deepEqual(bodyParagraphs(doc, content), [
        "İSTANBUL 3. ASLİYE HUKUK MAHKEMESİ'NE",
        'IŞIK İĞNE',
        'INDIGO INK',
        'düz metin'
    ]);
});

test('All Caps set by a heading style, and its language from the document defaults, are resolved', async () => {
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>' +
        paragraphStyle('Baslik1', '<w:caps/>') + characterStyle('Turkce', '<w:lang w:val="tr-TR"/>');
    const { doc, content } = await convertBody(
        p(r('mahkemesi'), '<w:pStyle w:val="Baslik1"/>') +
        p(r('mahkemesi', '<w:rStyle w:val="Turkce"/>'), '<w:pStyle w:val="Baslik1"/>') +
        p(r('mahkemesi', '<w:caps w:val="0"/>'), '<w:pStyle w:val="Baslik1"/>'), { styles });

    assert.deepEqual(bodyParagraphs(doc, content), ['MAHKEMESI', 'MAHKEMESİ', 'mahkemesi']);
});

test('Small Caps text is written in capitals', async () => {
    const { doc, content } = await convertBody(p(r('Davacı ') + r('vekili', '<w:smallCaps/>')));
    assert.deepEqual(bodyParagraphs(doc, content), ['Davacı VEKİLİ']);
});

test('All Caps applies in table cells and the header too', async () => {
    const { doc, content } = await convertBody(tbl([p(r('hücre', '<w:caps/>'))]) + p(r('Gövde')), {
        sectPr: headerRef,
        parts: [header(p(r('başlık', '<w:caps/>')))]
    });
    assert.deepEqual(cellParagraphs(doc, content), ['HÜCRE']);
    const headerEl = doc.querySelector('elements > header');
    assert.equal([...headerEl.querySelectorAll('[startOffset]')].map(el => textOf(el, content)).join(''), 'BAŞLIK');
});

test('a non-breaking hyphen between two w:t elements becomes a hyphen-minus; a soft hyphen stays dropped', async () => {
    const { doc, content, parsed } = await convertBody(
        p('<w:r><w:t>2024</w:t><w:noBreakHyphen/><w:t>123</w:t></w:r>') +
        p('<w:r><w:t>Esas </w:t></w:r><w:r><w:noBreakHyphen/></w:r><w:r><w:t>No</w:t></w:r>') +
        p('<w:r><w:t>hukuk</w:t><w:softHyphen/><w:t>çu</w:t></w:r>'));

    assert.deepEqual(bodyParagraphs(doc, content), ['2024-123', 'Esas -No', 'hukukçu']);
    // The hyphen joins its run's text instead of starting a new piece
    assert.deepEqual(plain(parsed.elements[0].runs.map(run => run.text)), ['2024-123']);
});
