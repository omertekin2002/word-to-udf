const { test } = require('node:test');
const assert = require('node:assert/strict');
const { convertBody, textOf } = require('./helpers');

// Word equations (Office Math, m:oMath): written as linear text and reported as equations

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const tbl = inner => '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`;
const textBox = inner => '<w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const footnoteRef = id => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const footnoteMark = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' + r(' ');

// Office Math building blocks, as Word writes them
const mr = text => `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>${text}</m:t></m:r>`;
const oMath = inner => `<m:oMath>${inner}</m:oMath>`;
const el = (name, inner) => `<m:${name}>${inner}</m:${name}>`;
const ctrlPr = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';
const frac = (num, den) => el('f', `<m:fPr>${ctrlPr}</m:fPr>` + el('num', num) + el('den', den));
const sup = (base, script) => el('sSup', `<m:sSupPr>${ctrlPr}</m:sSupPr>` + el('e', base) + el('sup', script));
const sub = (base, script) => el('sSub', el('e', base) + el('sub', script));
const subSup = (base, lower, upper) => el('sSubSup', el('e', base) + el('sub', lower) + el('sup', upper));
const rad = (degree, radicand) => el('rad', (degree === null ? '<m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>' : el('deg', degree)) +
    el('e', radicand));
const delim = (dPr, ...parts) => el('d', (dPr ? `<m:dPr>${dPr}</m:dPr>` : '') + parts.map(part => el('e', part)).join(''));
const chr = (name, value) => `<m:${name} m:val="${value}"/>`;
const nary = (naryPr, lower, upper, body) => el('nary', `<m:naryPr>${naryPr}${ctrlPr}</m:naryPr>` +
    el('sub', lower) + el('sup', upper) + el('e', body));
const func = (name, argument) => el('func', el('fName', name) + el('e', argument));

// The parser runs in its own realm; copy its values into this one so deepEqual can compare them
const plain = value => JSON.parse(JSON.stringify(value));
const finding = (parsed, code) => plain(parsed.findings).find(f => f.code === code);

test('an equation in a sentence is written as linear text where it stands, and reported with its place', async () => {
    const { content, parsed } = await convertBody(p(r('Faiz: ') + oMath(frac(mr('A'), mr('365'))) + r(' TL')));
    assert.equal(content, 'Faiz: A/365 TL');
    assert.deepEqual(finding(parsed, 'equations'), {
        code: 'equations', count: 1, examples: ['A/365'],
        occurrences: [{ example: 'A/365', location: { kind: 'paragraph', paragraph: 1 } }]
    });
});

const LINEAR_FORMS = [
    ['a fraction with a compound numerator', frac(mr('a+b'), mr('2')), '(a+b)/2'],
    ['a fraction whose numerator is already bracketed', frac(delim('', mr('a+b')), mr('2')), '(a+b)/2'],
    ['a fraction whose numerator is a delimiter without brackets',
        frac(delim(chr('begChr', '') + chr('endChr', ''), mr('a+b')), mr('2')), '(a+b)/2'],
    ['a fraction whose numerator is a delimiter with only a begin bracket',
        frac(delim(chr('begChr', '{') + chr('endChr', ''), mr('a+b')), mr('2')), '({a+b)/2'],
    ['a fraction of decimal numbers', frac(mr('1.5'), mr('365')), '1.5/365'],
    ['a superscript', sup(mr('x'), mr('2')), 'x^2'],
    ['a compound superscript', sup(mr('x'), mr('n+1')), 'x^(n+1)'],
    ['a subscript', sub(mr('x'), mr('i')), 'x_i'],
    ['a subscript and superscript', subSup(mr('x'), mr('i'), mr('2')), 'x_i^2'],
    ['a square root with its degree hidden', rad(null, mr('x')), '√(x)'],
    ['a square root with no degree element', el('rad', el('e', mr('x'))), '√(x)'],
    ['a cube root', rad(mr('3'), mr('x')), '3√(x)'],
    ['a delimiter with its own brackets and two parts', delim(chr('begChr', '[') + chr('endChr', ']'), mr('a'), mr('b')), '[a|b]'],
    ['a delimiter with the default brackets and a separator', delim(chr('sepChr', ';'), mr('a'), mr('b')), '(a;b)'],
    ['a delimiter with no end bracket', delim(chr('begChr', '{') + chr('endChr', ''), mr('x')), '{x'],
    ['a summation with both limits', nary(chr('chr', '∑'), mr('i=1'), mr('n'), mr('i')), '∑_(i=1)^n i'],
    ['an integral, the default operator, with its limits hidden',
        nary('<m:subHide m:val="on"/><m:supHide/>', '', '', mr('f(x)dx')), '∫ f(x)dx'],
    ['a function', func(mr('sin'), mr('x')), 'sin x'],
    ['an accent, by its text', el('acc', `<m:accPr>${chr('chr', '̂')}${ctrlPr}</m:accPr>` + el('e', mr('x'))), 'x'],
    ['a bar, by its text', el('bar', `<m:barPr><m:pos m:val="top"/>${ctrlPr}</m:barPr>` + el('e', mr('x+y'))), 'x+y'],
    ['a matrix with distinct rows and entries',
        el('m', el('mr', el('e', mr('1')) + el('e', mr('2'))) + el('mr', el('e', mr('3')) + el('e', mr('4')))),
        'matrix([1, 2], [3, 4])'],
    ['a binomial coefficient inside parentheses',
        delim('', el('f', '<m:fPr><m:type m:val="noBar"/></m:fPr>' + el('num', mr('5')) + el('den', mr('2')))),
        '(binom(5, 2))'],
    ['nested structures', frac(sup(mr('x'), mr('2')), func(mr('cos'), mr('y'))), '(x^2)/(cos y)']
];

for (const [what, math, linear] of LINEAR_FORMS) {
    test(`${what} is written as ${linear}`, async () => {
        const { content, parsed } = await convertBody(p(oMath(math)));
        assert.equal(content, linear);
        assert.deepEqual(finding(parsed, 'equations').examples, [linear]);
    });
}

test('matrix rows and cells can come from selected compatibility branches', async () => {
    const alternate = (choice, fallback, requires) => '<mc:AlternateContent>' +
        `<mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback>` +
        '</mc:AlternateContent>';
    const row = (...values) => el('mr', values.map(value => el('e', mr(value))).join(''));
    const math = el('m', alternate(row('ignored'), row('1', '2'), 'unsupported') +
        el('mr', alternate(el('e', mr('3')), el('e', mr('ignored')), 'w') + el('e', mr('4'))));
    const { content, parsed } = await convertBody(p(oMath(math)));
    assert.equal(content, 'matrix([1, 2], [3, 4])');
    assert.deepEqual(finding(parsed, 'equations').examples, [content]);
});

test('runs deleted inside an equation are left out, inserted ones kept', async () => {
    const math = oMath(mr('x') + `<w:del w:id="1" w:author="Y"><m:r><m:t>+1</m:t></m:r></w:del>` +
        `<w:ins w:id="2" w:author="Y"><m:r><m:t>-1</m:t></m:r></w:ins>`);
    const { content } = await convertBody(p(math));
    assert.equal(content, 'x-1');
});

test('an m:oMathPara writes each of its equations on its own line, and reports each one', async () => {
    const { content, parsed } = await convertBody(p(`<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>` +
        oMath(mr('a=b')) + oMath(sup(mr('c'), mr('2'))) + '</m:oMathPara>'));
    assert.equal(content, 'a=b\nc^2');
    assert.deepEqual(finding(parsed, 'equations'), {
        code: 'equations', count: 2, examples: ['a=b', 'c^2'],
        occurrences: [
            { example: 'a=b', location: { kind: 'paragraph', paragraph: 1 } },
            { example: 'c^2', location: { kind: 'paragraph', paragraph: 1 } }
        ]
    });
});

test('the text around an equation keeps its formatting; the equation takes the paragraph\'s, not Cambria Math', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Govde"><w:name w:val="Govde"/>' +
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:style>';
    const { doc, content } = await convertBody(
        p(r('Faiz: ', '<w:b/>') + oMath(frac(mr('A'), mr('365'))) + r(' TL', '<w:i/>'), '<w:pStyle w:val="Govde"/>'),
        { styles });
    assert.equal(content, 'Faiz: A/365 TL');
    const pieces = [...doc.querySelectorAll('elements > paragraph > content')]
        .map(piece => ({ text: textOf(piece, content), family: piece.getAttribute('family'), bold: piece.getAttribute('bold'),
            italic: piece.getAttribute('italic') }));
    assert.deepEqual(pieces, [
        { text: 'Faiz: ', family: 'Arial', bold: 'true', italic: null },
        { text: 'A/365', family: 'Arial', bold: null, italic: null },
        { text: ' TL', family: 'Arial', bold: null, italic: 'true' }
    ]);
});

test('a hidden equation is left out and reported as hidden text, not as an equation', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Gizli"><w:name w:val="Gizli"/><w:rPr><w:vanish/></w:rPr></w:style>';
    const { content, parsed } = await convertBody(p(r('Görünür') + oMath(mr('x')), '<w:pStyle w:val="Gizli"/>') + p(r('Son')),
        { styles });
    // The emptied paragraph is written as the generator writes any empty one
    assert.equal(content.replace(/\u200B/g, ''), 'Son');
    assert.equal(finding(parsed, 'equations'), undefined);
    assert.deepEqual(finding(parsed, 'hiddenText').examples, ['Görünürx']);
});

// An equation run with its own w:rPr
const mrWith = (text, rPr) => `<m:r><w:rPr>${rPr}</w:rPr><m:t>${text}</m:t></m:r>`;

test('a hidden run inside an equation is left out and reported as hidden text; the rest is written', async () => {
    const { content, parsed } = await convertBody(p(r('Faiz: ') + oMath(mr('A') + mrWith('+gizli', '<w:vanish/>') + mr('/365'))));
    assert.equal(content, 'Faiz: A/365');
    assert.deepEqual(finding(parsed, 'equations').examples, ['A/365']);
    assert.deepEqual(finding(parsed, 'hiddenText').examples, ['+gizli']);
});

test('a structure whose runs are all hidden leaves no operator behind', async () => {
    for (const math of [
        frac(mrWith('a', '<w:vanish/>'), mrWith('b', '<w:vanish/>')),
        sup(mrWith('x', '<w:vanish/>'), mrWith('2', '<w:vanish/>')),
        rad(null, mrWith('x', '<w:vanish/>'))
    ]) {
        const { content, parsed } = await convertBody(p(r('Metin') + oMath(math)));
        assert.equal(content, 'Metin');
        assert.equal(finding(parsed, 'equations'), undefined);
        assert.equal(finding(parsed, 'hiddenText').count, 1);
    }
    // The same in a footnote
    const { parsed } = await convertBody(p(r('Faiz') + footnoteRef(1)), {
        footnotes: { 1: p(footnoteMark + r('Oran') + oMath(frac(mrWith('a', '<w:vanish/>'), mrWith('b', '<w:vanish/>')))) }
    });
    assert.equal(parsed.footnotes[1], 'Oran');
    assert.equal(finding(parsed, 'equations'), undefined);
});

test('an equation run hidden through its character style is left out; w:vanish off on a run shows it again', async () => {
    const styles = '<w:style w:type="character" w:styleId="Gizli"><w:name w:val="Gizli"/><w:rPr><w:vanish/></w:rPr></w:style>';
    const { content, parsed } = await convertBody(
        p(oMath(frac(mrWith('x', '<w:rStyle w:val="Gizli"/>') + mr('a'), mrWith('2', '<w:rStyle w:val="Gizli"/><w:vanish w:val="0"/>')))),
        { styles });
    assert.equal(content, 'a/2');
    assert.deepEqual(finding(parsed, 'hiddenText').examples, ['x']);
});

test('an equation whose runs are all hidden is left out and reported only as hidden text', async () => {
    const { content, parsed } = await convertBody(p(r('Metin') + oMath(mrWith('x', '<w:vanish/>'))));
    assert.equal(content, 'Metin');
    assert.equal(finding(parsed, 'equations'), undefined);
    assert.deepEqual(finding(parsed, 'hiddenText').examples, ['x']);
});

test('a hidden run inside a footnote equation is left out of the note and reported at the reference', async () => {
    const { content, parsed } = await convertBody(p(r('Faiz') + footnoteRef(1)), {
        footnotes: { 1: p(footnoteMark + r('Oran ') + oMath(mr('A') + mrWith('gizli', '<w:vanish/>') + mr('/365'))) }
    });
    assert.ok(content.includes('Oran A/365') && !content.includes('gizli'), JSON.stringify(content));
    assert.equal(parsed.footnotes[1], 'Oran A/365');
    const inNote = { kind: 'paragraph', paragraph: 1, note: { kind: 'footnote' } };
    assert.deepEqual(finding(parsed, 'hiddenText').occurrences, [{ example: 'gizli', location: inNote }]);
    assert.deepEqual(finding(parsed, 'equations').occurrences, [{ example: 'A/365', location: inNote }]);
});

test('an empty equation placeholder writes nothing and is not reported', async () => {
    const { content, parsed } = await convertBody(p(r('Metin') + oMath(ctrlPr)));
    assert.equal(content, 'Metin');
    assert.deepEqual(plain(parsed.findings), []);
});

test('equations in a table cell, a text box and the default header are written and located there', async () => {
    const { content, parsed } = await convertBody(
        tbl(p(r('Hücre ') + oMath(sup(mr('x'), mr('2'))))) + p(r('Çapa') + textBox(p(oMath(sub(mr('y'), mr('j')))))),
        {
            sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>',
            parts: [{ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(r('Üst ') + oMath(mr('z'))) }]
        });
    for (const text of ['Hücre x^2', 'y_j', 'Üst z']) assert.ok(content.includes(text), `${text} in ${JSON.stringify(content)}`);
    assert.deepEqual(finding(parsed, 'equations').occurrences, [
        { example: 'z', location: { kind: 'header' } },
        { example: 'x^2', location: { kind: 'cell', table: 1, row: 1, cell: 1 } },
        { example: 'y_j', location: { kind: 'paragraph', paragraph: 1 } }
    ]);
});

test('a header whose only text is an equation is written', async () => {
    const { content, parsed } = await convertBody(p(r('Gövde')), {
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/>',
        parts: [{ name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1', inner: p(oMath(frac(mr('1'), mr('2')))) }]
    });
    assert.ok(content.includes('1/2'), JSON.stringify(content));
    assert.deepEqual(finding(parsed, 'equations').occurrences, [{ example: '1/2', location: { kind: 'header' } }]);
});

test('an equation in a footnote is written as linear text in the note, and reported at the reference', async () => {
    // Built: the note's own paragraphs carry the equation's linear text, and so does its plain text
    const { content, parsed } = await convertBody(p(r('Faiz') + footnoteRef(1)), {
        footnotes: { 1: p(footnoteMark + r('Oran ') + oMath(frac(mr('A'), mr('365'))) + r(' olarak')) }
    });
    assert.ok(content.includes('Oran A/365 olarak'), JSON.stringify(content));
    assert.equal(parsed.footnotes[1], 'Oran A/365 olarak');
    assert.deepEqual(finding(parsed, 'equations').occurrences, [
        { example: 'A/365', location: { kind: 'paragraph', paragraph: 1, note: { kind: 'footnote' } } }
    ]);
});
