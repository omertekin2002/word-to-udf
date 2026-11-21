const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf, findingSummaries } = require('./helpers');

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraphStyle = (id, rPr) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/><w:rPr>${rPr}</w:rPr></w:style>`;
const footnoteRef = id => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const endnoteRef = id => `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
// The first paragraph of a note as Word writes it: the mark run, then a space before the text
const footnoteMark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' + r(' ');
const endnoteMark = '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>' + r(' ');
const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
const SEPARATOR = '────────────────────';

// Java-style signed ARGB integer of a hex color, as the generator writes it
const argb = hex => (0xFF000000 | parseInt(hex, 16)) | 0;

/**
 * The paragraphs of every notes section, in order: each paragraph's elements with their text
 * and attributes. A notes section starts after its separator line.
 */
function notesParagraphs(doc, content) {
    const paragraphs = [...doc.querySelectorAll('elements > paragraph')];
    const result = [];
    let inNotes = false;
    for (const paragraph of paragraphs) {
        const pieces = [...paragraph.querySelectorAll('[startOffset]')];
        const text = pieces.map(piece => textOf(piece, content)).join('');
        if (text === SEPARATOR || text === '\n') {
            // The separator starts a notes section; the empty line before one ends the section above
            inNotes = text === SEPARATOR;
            continue;
        }
        if (!inNotes) continue;
        result.push({ element: paragraph, text, pieces: pieces.map(piece => ({ piece, text: textOf(piece, content) })) });
    }
    return result;
}

/**
 * Every element that covers content does so without gaps or overlaps: in document order, each
 * starts where the one before it ends, and the last ends at the end of the content. Holds for
 * documents without tables (a cell adds a newline between its paragraphs that no element covers).
 */
function assertContentCovered({ doc, content }) {
    let cursor = 0;
    for (const el of doc.querySelectorAll('elements [startOffset]')) {
        const start = Number(el.getAttribute('startOffset'));
        assert.equal(start, cursor, `<${el.localName}> starts at ${start}, expected ${cursor}`);
        cursor = start + Number(el.getAttribute('length'));
    }
    assert.equal(cursor, content.length, 'the elements cover the whole content');
}

const attributes = piece => Object.fromEntries([...piece.attributes]
    .filter(attribute => !['startOffset', 'length'].includes(attribute.name))
    .map(attribute => [attribute.name, attribute.value]));

test('a footnote keeps bold, italic, underline, color and highlighting on its runs, with its text exactly as in Word', async () => {
    const result = await convertBody(p(r('Dilekçe') + footnoteRef(1)), {
        footnotes: {
            1: p(footnoteMark + r('Yargıtay 3. HD', '<w:b/>') + r(', ') + r('Emsal', '<w:i/>') + r(' ve ') +
                r('altı çizili', '<w:u w:val="single"/>') + r(' / ') + r('kırmızı', '<w:color w:val="FF0000"/>') +
                r(' / ') + r('vurgulu', '<w:highlight w:val="yellow"/>'))
        }
    });
    const { doc, content, parsed } = result;
    assertContentCovered(result);

    const [note] = notesParagraphs(doc, content);
    assert.equal(note.text, '1. Yargıtay 3. HD, Emsal ve altı çizili / kırmızı / vurgulu');
    const byText = Object.fromEntries(note.pieces.map(({ piece, text }) => [text, attributes(piece)]));
    assert.deepEqual(byText['1. '], { family: 'Times New Roman', size: '10', superscript: 'true' });
    assert.deepEqual(byText['Yargıtay 3. HD'], { family: 'Times New Roman', size: '10', bold: 'true' });
    assert.deepEqual(byText['Emsal'], { family: 'Times New Roman', size: '10', italic: 'true' });
    assert.deepEqual(byText['altı çizili'], { family: 'Times New Roman', size: '10', underline: 'true' });
    assert.deepEqual(byText['kırmızı'], { family: 'Times New Roman', size: '10', foreground: String(argb('FF0000')) });
    assert.deepEqual(byText['vurgulu'], { family: 'Times New Roman', size: '10', background: String(argb('FFFF00')) });
    assert.deepEqual(byText[', '], { family: 'Times New Roman', size: '10' });
    // The plain text is still the note's characters, without the space after the mark
    assert.equal(parsed.footnotes[1], 'Yargıtay 3. HD, Emsal ve altı çizili / kırmızı / vurgulu');
});

test('a two-paragraph footnote is written as two paragraphs; only the first starts with the number', async () => {
    const result = await convertBody(p(r('Metin') + footnoteRef(1)) + p(r('Sonraki')), {
        footnotes: { 1: p(footnoteMark + r('Yargıtay 3. HD, E. 2020/1', '<w:b/>')) + p(r('Açıklama: ') + r('kesinleşmiştir', '<w:i/>')) }
    });
    const { doc, content } = result;
    assertContentCovered(result);

    const notes = notesParagraphs(doc, content);
    assert.deepEqual(notes.map(note => note.text), ['1. Yargıtay 3. HD, E. 2020/1', 'Açıklama: kesinleşmiştir']);
    assert.equal(notes[0].pieces[0].piece.getAttribute('superscript'), 'true');
    assert.equal(notes[1].element.querySelector('[superscript]'), null);
    assert.equal(notes[1].pieces[1].piece.getAttribute('italic'), 'true');
    for (const note of notes) {
        assert.equal(note.element.getAttribute('Alignment'), '0');
        assert.equal(note.element.getAttribute('LeftIndent'), '0.0');
        assert.equal(note.element.getAttribute('RightIndent'), '0.0');
    }
    assert.ok(content.endsWith('1. Yargıtay 3. HD, E. 2020/1Açıklama: kesinleşmiştir'));
});

test('a line break in a note is a break and a tab is a <tab>; the space or tab after the mark is dropped', async () => {
    for (const afterMark of [' ', '  ', '<w:tab/>']) {
        const mark = '<w:r><w:footnoteRef/></w:r>' + (afterMark.startsWith('<') ? `<w:r>${afterMark}</w:r>` : r(afterMark));
        const result = await convertBody(p(r('A') + footnoteRef(1)), {
            footnotes: { 1: p(mark + r('Karar') + '<w:r><w:br/></w:r>' + r('E. 2020/1') + '<w:r><w:tab/></w:r>' + r('K. 2021/2')) }
        });
        const { doc, content } = result;
        assertContentCovered(result);

        const [note] = notesParagraphs(doc, content);
        assert.equal(note.text, '1. Karar\nE. 2020/1\tK. 2021/2', JSON.stringify(afterMark));
        const tab = note.pieces.find(({ text }) => text === '\t');
        assert.equal(tab.piece.localName, 'tab');
        const newline = note.pieces.find(({ text }) => text === '\n');
        assert.equal(newline.piece.localName, 'content');
    }
});

test('only one tab after the mark is dropped, and spaces later in the note stay', async () => {
    const { doc, content } = await convertBody(p(r('A') + footnoteRef(1)), {
        footnotes: { 1: p('<w:r><w:footnoteRef/></w:r><w:r><w:tab/><w:tab/></w:r>' + r('Metin  iki')) }
    });
    assert.equal(notesParagraphs(doc, content)[0].text, '1. \tMetin  iki');
});

test('formatting set on the footnote text style reaches the note runs, and direct formatting overrides it', async () => {
    const styles = paragraphStyle('FootnoteText', '<w:i/><w:sz w:val="18"/>');
    const result = await convertBody(p(r('A') + footnoteRef(1)), {
        styles,
        footnotes: { 1: p(footnoteMark + r('Stilden') + r(' büyük', '<w:i w:val="0"/><w:sz w:val="22"/>'), '<w:pStyle w:val="FootnoteText"/>') }
    });
    const { doc, content } = result;
    assertContentCovered(result);

    const [note] = notesParagraphs(doc, content);
    const byText = Object.fromEntries(note.pieces.map(({ piece, text }) => [text, attributes(piece)]));
    assert.deepEqual(byText['Stilden'], { family: 'Times New Roman', size: '9', italic: 'true' });
    assert.deepEqual(byText[' büyük'], { family: 'Times New Roman', size: '11' });
});

test('a note run without a size anywhere is 10 pt; a size in the document defaults or a font applies as in the body', async () => {
    const plainNote = { 1: p(footnoteMark + r('Kaynak')) };
    const sizeOf = async options => {
        const { doc, content } = await convertBody(p(r('A') + footnoteRef(1)), { footnotes: plainNote, ...options });
        return attributes(notesParagraphs(doc, content)[0].pieces[1].piece);
    };

    assert.deepEqual(await sizeOf({}), { family: 'Times New Roman', size: '10' });
    assert.deepEqual(await sizeOf({ styles: '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' }),
        { family: 'Times New Roman', size: '12' });
    assert.deepEqual(await sizeOf({ styles: '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr></w:rPrDefault></w:docDefaults>' }),
        { family: 'Arial', size: '10' });
});

test('endnotes get the same treatment, keep their number format and still go last; footnotes still flush at a page break', async () => {
    const result = await convertBody(
        p(r('Bir') + footnoteRef(1) + endnoteRef(1) + PAGE_BREAK + r('İki') + footnoteRef(2) + endnoteRef(2)), {
            sectPr: '<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>',
            footnotes: {
                1: p(footnoteMark + r('Dipnot bir', '<w:b/>')) + p(r('ikinci paragraf')),
                2: p(footnoteMark + r('Dipnot iki', '<w:i/>'))
            },
            endnotes: {
                1: p(endnoteMark + r('Son not', '<w:u w:val="single"/>') + '<w:r><w:br/></w:r>' + r('devamı')),
                2: p(endnoteMark + r('İkinci son not')) + p(r('ek', '<w:b/>'))
            }
        });
    const { doc, content } = result;
    assertContentCovered(result);

    const children = [...doc.querySelector('elements').children].map(el => el.localName === 'page-break'
        ? 'page-break'
        : [...el.querySelectorAll('[startOffset]')].map(piece => textOf(piece, content)).join(''));
    assert.deepEqual(children, [
        'Bir1I',
        '\n', SEPARATOR, '1. Dipnot bir', 'ikinci paragraf',
        'page-break',
        'İki2II',
        '\n', SEPARATOR, '2. Dipnot iki',
        '\n', SEPARATOR, 'I. Son not\ndevamı', 'II. İkinci son not', 'ek'
    ]);

    const notes = notesParagraphs(doc, content);
    const piece = text => notes.flatMap(note => note.pieces).find(entry => entry.text === text).piece;
    assert.equal(piece('Dipnot bir').getAttribute('bold'), 'true');
    assert.equal(piece('Dipnot iki').getAttribute('italic'), 'true');
    assert.equal(piece('Son not').getAttribute('underline'), 'true');
    assert.equal(piece('ek').getAttribute('bold'), 'true');
    assert.equal(piece('I. ').getAttribute('superscript'), 'true');
});

test('hidden text, All Caps and symbols in a note behave as in the body, with the findings recorded per reference', async () => {
    const result = await convertBody(p(r('Metin') + footnoteRef(1) + r(' tekrar') + footnoteRef(1)), {
        footnotes: {
            1: p(footnoteMark + r('gizli not', '<w:vanish/>') + r('istinaf', '<w:caps/>') + r(' ') +
                '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' + '<w:r><w:noBreakHyphen/></w:r>' + r('son'))
        }
    });
    const { doc, content, parsed } = result;
    assertContentCovered(result);

    const notes = notesParagraphs(doc, content);
    assert.deepEqual(notes.map(note => note.text), ['1. İSTİNAF ✓-son', '2. İSTİNAF ✓-son']);
    assert.equal(content.includes('gizli'), false);
    // The plain text, used for the notesMoved examples, is the note's characters as before
    assert.equal(parsed.footnotes[1], 'istinaf ✓son');
    assert.deepEqual(findingSummaries(parsed.findings), [
        { code: 'hiddenText', count: 2, examples: ['gizli not'] },
        { code: 'notesMoved', count: 2, examples: ['istinaf ✓son'] },
        { code: 'symbols', count: 2, examples: ['Wingdings F0FC → ✓'] }
    ]);
    const occurrences = JSON.parse(JSON.stringify(parsed.findings)).find(finding => finding.code === 'symbols').occurrences;
    assert.deepEqual(occurrences.map(occurrence => occurrence.location.note), [{ kind: 'footnote' }]);
});

test('pictures, page breaks and note references inside a note are left out, and add no finding', async () => {
    const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData>' +
        '<pic:pic><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const result = await convertBody(p(r('A') + footnoteRef(1) + endnoteRef(1)), {
        footnotes: { 1: p(footnoteMark + r('Önce') + drawing + PAGE_BREAK + r('sonra')) },
        endnotes: { 1: p(endnoteMark + r('Son') + footnoteRef(1) + r(' not')) }
    });
    const { doc, content, parsed } = result;
    assertContentCovered(result);

    assert.deepEqual(notesParagraphs(doc, content).map(note => note.text), ['1. Öncesonra', 'i. Son not']);
    assert.equal(doc.querySelectorAll('image').length, 0);
    assert.equal(doc.querySelectorAll('page-break').length, 0);
    assert.deepEqual(findingSummaries(parsed.findings).map(finding => finding.code), ['notesMoved']);
});

test('an empty note keeps its placeholder, and an empty further paragraph is still a paragraph', async () => {
    const result = await convertBody(p(r('A') + footnoteRef(1) + endnoteRef(1) + footnoteRef(2)), {
        footnotes: { 1: p(footnoteMark), 2: p(footnoteMark + r('Birinci')) + p('') + p(r('Üçüncü')) },
        endnotes: { 1: p('<w:r><w:endnoteRef/></w:r><w:r><w:tab/></w:r>') }
    });
    const { doc, content } = result;
    assertContentCovered(result);

    assert.deepEqual(notesParagraphs(doc, content).map(note => note.text),
        ['1. [Empty footnote]', '2. Birinci', '​', 'Üçüncü', 'i. [Empty endnote]']);
});
