const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');
const { petitions } = require('./fixtures/petitions');

// Double strikethrough (w:dstrike). Docs.md gives no double strikethrough, so double-struck text is
// written with strikethrough="true", as single strikethrough (w:strike) is, and reported as a
// doubleStrikethrough conversion note.

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraphStyle = (id, rPr, basedOn = '') => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>` +
    `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}<w:rPr>${rPr}</w:rPr></w:style>`;
const characterStyle = (id, rPr, basedOn = '') => `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${id}"/>` +
    `${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}<w:rPr>${rPr}</w:rPr></w:style>`;
const tbl = cells => '<w:tbl><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="4000"/>').join('') + '</w:tblGrid>' +
    `<w:tr>${cells.map(inner => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`).join('')}</w:tr></w:tbl>`;
const footnoteRef = id => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const HEADING_STYLE = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>';

const DSTRIKE = '<w:dstrike/>';

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);
const occurrences = (parsed, code) => (finding(parsed, code) || { occurrences: [] }).occurrences;

// The attributes of every <content> whose text is text (offsets left out)
const attrsOf = (doc, content, text) => [...doc.querySelectorAll('content')]
    .filter(el => textOf(el, content) === text)
    .map(el => Object.fromEntries([...el.attributes]
        .filter(a => !['startOffset', 'length'].includes(a.name)).map(a => [a.name, a.value])));
const attrs = (doc, content, text) => {
    const all = attrsOf(doc, content, text);
    assert.equal(all.length, 1, `one <content> for ${text}`);
    return all[0];
};

const PLAIN = { family: 'Times New Roman', size: '12' };

test('the issue\'s reproduction: a direct w:dstrike run is written struck through, like w:strike, and reported', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('tek ', '<w:strike/>') + r('çift', DSTRIKE)) + p(r('Düz paragraf')));
    assert.deepEqual(attrs(doc, content, 'tek '), { ...PLAIN, strikethrough: 'true' });
    assert.deepEqual(attrs(doc, content, 'çift'), { ...PLAIN, strikethrough: 'true' });
    assert.deepEqual(attrs(doc, content, 'Düz paragraf'), PLAIN);
    assert.deepEqual(findingSummaries(parsed.findings), [{ code: 'doubleStrikethrough', count: 1, examples: ['tek çift'] }]);
    assert.deepEqual(occurrences(parsed, 'doubleStrikethrough'),
        [{ example: 'tek çift', location: { kind: 'paragraph', paragraph: 1 } }]);
});

test('a double-struck run keeps its other formatting, and no internal key reaches the parsed runs', async () => {
    const { doc, content, parsed } = await convertBody(
        p(r('Tutar: ') + r('1.000 TL', `<w:b/><w:i/><w:color w:val="FF0000"/><w:rFonts w:ascii="Arial"/><w:sz w:val="28"/>${DSTRIKE}`)));
    assert.deepEqual(attrs(doc, content, '1.000 TL'),
        { family: 'Arial', size: '14', bold: 'true', italic: 'true', strikethrough: 'true', foreground: '-65536' });
    for (const run of plain(parsed.elements[0].runs)) {
        for (const key of ['doubleStrike', 'dstrike', 'hidden', 'caps', 'smallCaps', 'lang']) assert.equal(key in run, false, key);
    }
});

test('a run with both w:strike and w:dstrike writes one strikethrough="true"', async () => {
    const { doc, xml, content, parsed } = await convertBody(p(r('ikisi', `<w:strike/>${DSTRIKE}`)));
    assert.deepEqual(attrs(doc, content, 'ikisi'), { ...PLAIN, strikethrough: 'true' });
    assert.doesNotMatch(xml, /strikethrough="true"[^>]*strikethrough=/);
    assert.equal(finding(parsed, 'doubleStrikethrough').count, 1);
});

test('double strikethrough from a paragraph style, a character style (and its base), or the document defaults is written', async () => {
    const styles = characterStyle('Cift', DSTRIKE) + characterStyle('CiftAlt', '', 'Cift') +
        paragraphStyle('CiftPara', DSTRIKE);
    const { doc, content, parsed } = await convertBody(
        p(r('düz') + r('karakter', '<w:rStyle w:val="Cift"/>') + r('zincir', '<w:rStyle w:val="CiftAlt"/>')) +
        p(r('paragraf'), '<w:pStyle w:val="CiftPara"/>'), { styles });
    assert.deepEqual(attrs(doc, content, 'düz'), PLAIN);
    for (const text of ['karakter', 'zincir', 'paragraf']) assert.equal(attrs(doc, content, text).strikethrough, 'true', text);
    assert.equal(finding(parsed, 'doubleStrikethrough').count, 2);

    const defaults = `<w:docDefaults><w:rPrDefault><w:rPr>${DSTRIKE}</w:rPr></w:rPrDefault></w:docDefaults>`;
    const fromDefaults = await convertBody(p(r('varsayılan')), { styles: defaults });
    assert.equal(attrs(fromDefaults.doc, fromDefaults.content, 'varsayılan').strikethrough, 'true');
});

test('w:dstrike w:val="0" (or false, off) on a run turns an inherited one off: no attribute, no note', async () => {
    const styles = characterStyle('Cift', DSTRIKE) + paragraphStyle('CiftPara', DSTRIKE);
    const { doc, content, parsed } = await convertBody(
        p(r('sıfır', '<w:dstrike w:val="0"/>') + r(' yanlış', '<w:dstrike w:val="false"/>'), '<w:pStyle w:val="CiftPara"/>') +
        p(r('kapalı', '<w:rStyle w:val="Cift"/><w:dstrike w:val="off"/>')), { styles });
    for (const text of ['sıfır', ' yanlış', 'kapalı']) assert.deepEqual(attrs(doc, content, text), PLAIN, text);
    assert.equal(finding(parsed, 'doubleStrikethrough'), undefined);
});

test('double strikethrough is not a toggle property: on in both the paragraph and character style stays on', async () => {
    // Unlike w:strike, which two style levels toggle off (ECMA-376 Part 1, 17.7.3)
    const styles = characterStyle('Cift', DSTRIKE) + paragraphStyle('CiftPara', DSTRIKE) +
        characterStyle('Tek', '<w:strike/>') + paragraphStyle('TekPara', '<w:strike/>');
    const { doc, content } = await convertBody(
        p(r('çift', '<w:rStyle w:val="Cift"/>'), '<w:pStyle w:val="CiftPara"/>') +
        p(r('tek', '<w:rStyle w:val="Tek"/>'), '<w:pStyle w:val="TekPara"/>'), { styles });
    assert.equal(attrs(doc, content, 'çift').strikethrough, 'true');
    assert.equal(attrs(doc, content, 'tek').strikethrough, undefined);
});

test('a w:dstrike w:val="0" in a style overrides one from an earlier layer', async () => {
    const styles = paragraphStyle('CiftPara', DSTRIKE) + characterStyle('Kapali', '<w:dstrike w:val="0"/>');
    const { doc, content, parsed } = await convertBody(
        p(r('kapalı', '<w:rStyle w:val="Kapali"/>') + r(' açık'), '<w:pStyle w:val="CiftPara"/>'), { styles });
    assert.equal(attrs(doc, content, 'kapalı').strikethrough, undefined);
    assert.equal(attrs(doc, content, ' açık').strikethrough, 'true');
    assert.equal(finding(parsed, 'doubleStrikethrough').count, 1);
});

test('a hidden or whitespace-only double-struck run adds no note', async () => {
    const { content, parsed } = await convertBody(
        p(r('Görünen ') + r('gizli', `<w:vanish/>${DSTRIKE}`) + r(' ', DSTRIKE)) + p(r('\t', DSTRIKE)));
    assert.equal(content.includes('gizli'), false);
    assert.equal(finding(parsed, 'doubleStrikethrough'), undefined);
    assert.equal(finding(parsed, 'hiddenText').count, 1);
});

test('plain w:strike behaves as before and adds no note', async () => {
    const { doc, content, parsed } = await convertBody(p(r('çizili', '<w:strike/>') + r(' düz')));
    assert.deepEqual(attrs(doc, content, 'çizili'), { ...PLAIN, strikethrough: 'true' });
    assert.deepEqual(attrs(doc, content, ' düz'), PLAIN);
    assert.deepEqual(findingSummaries(parsed.findings), []);
});

test('double strikethrough is written and located in a table cell, the header and footer, and a footnote', async () => {
    const parts = [
        { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Büro ') + r('eski', DSTRIKE)) },
        { name: 'footer1.xml', root: 'ftr', type: 'footer', id: 'rIdFooter1', inner: p(r('Sayfa ') + r('silindi', DSTRIKE)) }
    ];
    const sectPr = '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>';
    const { doc, content, parsed } = await convertBody(
        p(r('DELİLLER'), '<w:pStyle w:val="Heading1"/>') +
        tbl([p(r('A')), p(r('Bedel ') + r('500 TL', DSTRIKE))]) +
        p(r('Dipnot') + footnoteRef(1)),
        { styles: HEADING_STYLE, parts, sectPr, footnotes: { 1: p(r('Kaynak ') + r('geri alındı', DSTRIKE)) } });

    assert.equal(attrs(doc, content, '500 TL').strikethrough, 'true');
    const header = [...doc.querySelectorAll('header content')].find(el => textOf(el, content) === 'eski');
    assert.equal(header.getAttribute('strikethrough'), 'true');
    const footer = [...doc.querySelectorAll('footer content')].find(el => textOf(el, content) === 'silindi');
    assert.equal(footer.getAttribute('strikethrough'), 'true');
    // The footnote's text keeps its run formatting
    const note = [...doc.querySelectorAll('content')].find(el => textOf(el, content) === 'geri alındı');
    assert.equal(note.getAttribute('strikethrough'), 'true');

    const found = occurrences(parsed, 'doubleStrikethrough');
    assert.deepEqual(found.map(o => o.example).sort(), ['Bedel 500 TL', 'Büro eski', 'Kaynak geri alındı', 'Sayfa silindi'].sort());
    assert.deepEqual(found.find(o => o.example === 'Bedel 500 TL').location,
        { kind: 'cell', table: 1, row: 1, cell: 2, heading: 'DELİLLER' });
    assert.deepEqual(found.find(o => o.example === 'Kaynak geri alındı').location,
        { kind: 'paragraph', paragraph: 2, heading: 'DELİLLER', note: { kind: 'footnote' } });
    assert.equal(finding(parsed, 'doubleStrikethrough').count, 4);
});

test('a hidden double-struck run in a footnote adds no note', async () => {
    const { parsed } = await convertBody(p(r('Dipnot') + footnoteRef(1)),
        { footnotes: { 1: p(r('Görünen') + r(' gizli', `<w:vanish/>${DSTRIKE}`) + r(' ', DSTRIKE)) } });
    assert.equal(finding(parsed, 'doubleStrikethrough'), undefined);
});

test('an equation in a double-struck paragraph style is written struck through and reported, in the body and a footnote', async () => {
    const oMath = text => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`;
    const styles = paragraphStyle('CiftPara', DSTRIKE);
    const { doc, content, parsed } = await convertBody(
        p(oMath('A/365'), '<w:pStyle w:val="CiftPara"/>') + p(oMath('x+1')) + p(r('Dipnot') + footnoteRef(1)),
        { styles, footnotes: { 1: p(oMath('B/2'), '<w:pStyle w:val="CiftPara"/>') } });
    assert.equal(attrs(doc, content, 'A/365').strikethrough, 'true');
    assert.equal(attrs(doc, content, 'x+1').strikethrough, undefined);
    assert.deepEqual(occurrences(parsed, 'doubleStrikethrough').map(o => o.example), ['A/365', 'B/2']);
});

test('the petition fixtures give no doubleStrikethrough note', async () => {
    for (const petition of petitions) {
        const { DocxParser } = require('./helpers').loadConverter();
        const parsed = await new DocxParser().parse(await petition.build());
        assert.ok(!plain(parsed.findings).some(f => f.code === 'doubleStrikethrough'), petition.name);
    }
});
