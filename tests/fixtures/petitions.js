/**
 * Synthetic petition fixtures: whole documents shaped like typical Turkish court
 * filings, built in code (no binary .docx is committed).
 *
 * Every name, ID number, case number and address here is an invented placeholder.
 * Never put real personal data in these fixtures.
 *
 * Each fixture records its text runs while it is built, so the tests know which
 * strings the UDF must contain and in what order.
 */

const { buildDocx } = require('../helpers');

const escapeText = text => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * Collects a fixture's text as its XML is built, in document order.
 * - texts: body text pieces as the UDF content holds them, including the number
 *   a footnote reference becomes
 * - footnotes: footnote texts in reference order (the UDF appends them after the body)
 * - deleted: tracked-deletion text that must not appear
 */
function recorder() {
    const texts = [];
    const footnotes = [];
    const footnoteXml = {};
    const deleted = [];

    const rPrXml = rPr => (rPr ? `<w:rPr>${rPr}</w:rPr>` : '');

    return {
        texts,
        footnotes,
        footnoteXml,
        deleted,

        /** A run holding one or more texts separated by tabs/breaks given as XML, e.g. run('A', '<w:br/>', 'B') */
        run(...pieces) {
            let rPr = '';
            if (pieces.length > 1 && typeof pieces[0] === 'object') rPr = pieces.shift().rPr;
            const inner = pieces.map(piece => {
                if (piece.startsWith('<')) return piece;
                texts.push(piece);
                return `<w:t xml:space="preserve">${escapeText(piece)}</w:t>`;
            }).join('');
            return `<w:r>${rPrXml(rPr)}${inner}</w:r>`;
        },

        /** A footnote reference run; the footnote body is a single paragraph */
        footnote(text) {
            const id = String(footnotes.length + 1);
            footnotes.push(text);
            texts.push(id);
            // Word's footnote paragraph: the reference mark, a space, then the text
            footnoteXml[id] = '<w:p><w:pPr><w:pStyle w:val="DipnotMetni"/></w:pPr>' +
                '<w:r><w:rPr><w:rStyle w:val="DipnotBavurusu"/></w:rPr><w:footnoteRef/></w:r>' +
                '<w:r><w:t xml:space="preserve"> </w:t></w:r>' +
                `<w:r><w:t>${escapeText(text)}</w:t></w:r></w:p>`;
            return `<w:r><w:rPr><w:rStyle w:val="DipnotBavurusu"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
        },

        /** A tracked deletion */
        del(text) {
            deleted.push(text);
            return '<w:del w:id="90" w:author="Yazar" w:date="2026-01-01T00:00:00Z">' +
                `<w:r><w:delText xml:space="preserve">${escapeText(text)}</w:delText></w:r></w:del>`;
        }
    };
}

/** The recorded strings a test checks against */
const recorded = ({ texts, footnotes, deleted }) => ({ texts, footnotes, deleted });

const p = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const empty = () => '<w:p/>';
const jc = val => `<w:jc w:val="${val}"/>`;
const list = (numId, level = 0) =>
    `<w:pStyle w:val="ListeParagraf"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
const BOLD = { rPr: '<w:b/>' };

// Tab stops for a party block: label, colon, value
const PARTY_TABS = '<w:tabs><w:tab w:val="left" w:pos="1985"/><w:tab w:val="left" w:pos="2268"/></w:tabs>' +
    '<w:ind w:left="2268" w:hanging="2268"/>';

/** A party-block line: LABEL<tab>:<tab>value, with optional further lines after line breaks */
function party(rec, label, ...values) {
    const lines = values.flatMap((value, i) => (i === 0 ? [value] : ['<w:br/>', value]));
    return p(rec.run(BOLD, label) + rec.run('<w:tab/>', ':', '<w:tab/>') + rec.run(...lines), PARTY_TABS);
}

// Word's default styles, with the names Turkish Word writes (Balk1 is "Başlık 1")
const BASE_STYLES =
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="tr-TR"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="VarsaylanParagrafYazTipi"><w:name w:val="Default Paragraph Font"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListeParagraf"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="DipnotMetni"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>' +
    '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="DipnotBavurusu"><w:name w:val="footnote reference"/>' +
    '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';

const HEADING_STYLE =
    '<w:style w:type="paragraph" w:styleId="Balk1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
    '<w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
    '<w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>';

/** One w:lvl of an abstract numbering definition */
const level = (ilvl, numFmt, lvlText, left) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/><w:lvlText w:val="${lvlText}"/>` +
    `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
const bulletLevel = (ilvl, left) =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>` +
    `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>` +
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>';
const abstractNum = (id, levels) =>
    `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`;
const num = (numId, abstractNumId) => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractNumId}"/></w:num>`;

/** A4 portrait page with the given margins in twips */
const a4 = ({ top, right, bottom, left }) =>
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}" w:header="709" w:footer="709" w:gutter="0"/>` +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/>';

/**
 * 1. Dava dilekçesi: centred bold court heading, a tab-aligned party block, a
 * decimal list whose items carry footnotes, a bulleted evidence list and a
 * right-aligned signature block.
 */
function davaDilekcesi() {
    const rec = recorder();
    const margins = { top: 1418, right: 1134, bottom: 1418, left: 1701 };
    // Word numbers lists in the order they were created: the bullets came first here
    const BULLETS = 1;
    const NUMBERS = 2;

    const heading = "NÖBETÇİ ASLİYE HUKUK MAHKEMESİ'NE";
    const body = [
        p(rec.run({ rPr: '<w:b/><w:sz w:val="24"/>' }, heading), jc('center')),
        empty(),
        party(rec, 'DAVACI', 'Ad Soyad (T.C. Kimlik No: 00000000000)', 'Örnek Mahallesi, Örnek Sokak No: 0, İlçe / İl'),
        party(rec, 'VEKİLİ', 'Av. Ad Soyad', 'Örnek Hukuk Bürosu, Örnek Caddesi No: 0, İlçe / İl'),
        party(rec, 'DAVALI', 'Örnek Ticaret Ltd. Şti.', 'Örnek Mahallesi, Güneş Sokak No: 0, İlçe / İl'),
        party(rec, 'KONU', 'Ödenmeyen hizmet bedelinin faiziyle birlikte tahsili talebidir.'),
        empty(),
        // Word often splits a word across runs (here by a revision session)
        p('<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>' +
            rec.run(BOLD, 'AÇIKLA') + '<w:proofErr w:type="spellStart"/>' + rec.run(BOLD, 'MALAR') +
            '<w:proofErr w:type="spellEnd"/>', jc('center')),
        p(rec.run('Müvekkil ile davalı arasında 01.01.2026 tarihli hizmet sözleşmesi imzalanmıştır.') +
            rec.footnote('Ek-1: Hizmet sözleşmesinin örneği.'), list(NUMBERS) + jc('both')),
        p(rec.run('Davalı, sözleşmeden doğan ödeme yükümlülüğünü yerine getirmemiş; ihtarnameye rağmen borcunu ödememiştir.') +
            rec.footnote('Ek-2: Örnek Noterliğinin 00.00.2026 tarihli ve 00000 yevmiye numaralı ihtarnamesi.'),
        list(NUMBERS) + jc('both')),
        p(rec.run('Bu nedenle işbu davayı açma zorunluluğu doğmuştur.'), list(NUMBERS) + jc('both')),
        empty(),
        p(rec.run(BOLD, 'HUKUKİ SEBEPLER'), jc('center')),
        p(rec.run('6098 sayılı Türk Borçlar Kanunu, 6100 sayılı HMK ve ilgili mevzuat.'), jc('both')),
        empty(),
        p(rec.run(BOLD, 'DELİLLER'), jc('center')),
        p(rec.run('Hizmet sözleşmesi'), list(BULLETS)),
        p(rec.run('İhtarname ve tebliğ şerhi'), list(BULLETS)),
        p(rec.run('Banka kayıtları ve faturalar'), list(BULLETS)),
        p(rec.run('Tanık beyanları, bilirkişi incelemesi ve her türlü yasal delil'), list(BULLETS)),
        empty(),
        p(rec.run(BOLD, 'SONUÇ VE İSTEM'), jc('center')),
        p(rec.run('Yukarıda açıklanan nedenlerle; davanın kabulüne, alacağın dava tarihinden itibaren işleyecek ' +
            'yasal faiziyle birlikte davalıdan tahsiline, yargılama giderleri ile vekâlet ücretinin davalıya ' +
            'yükletilmesine karar verilmesini saygılarımızla vekâleten arz ve talep ederiz.'), jc('both')),
        empty(),
        p(rec.run('Davacı Vekili'), jc('right')),
        p(rec.run('Av. Ad Soyad'), jc('right')),
        p(rec.run({ rPr: '<w:i/>' }, '(e-imzalıdır)'), jc('right'))
    ].join('');

    const numbering = abstractNum(0, [0, 1, 2].map(i => bulletLevel(i, 720 + i * 720)).join('')) +
        abstractNum(1, [0, 1, 2].map(i => level(i, 'decimal', `%${i + 1}.`, 720 + i * 720)).join('')) +
        num(BULLETS, 0) + num(NUMBERS, 1);

    return {
        name: 'dava dilekçesi',
        options: { styles: BASE_STYLES, numbering, footnotes: rec.footnoteXml, sectPr: a4(margins) },
        body,
        ...recorded(rec),
        margins,
        heading,
        listItems: [
            { text: 'Müvekkil ile davalı arasında 01.01.2026 tarihli hizmet sözleşmesi imzalanmıştır.', kind: 'number', level: 0 },
            { text: 'Davalı, sözleşmeden doğan ödeme yükümlülüğünü yerine getirmemiş; ihtarnameye rağmen borcunu ödememiştir.', kind: 'number', level: 0 },
            { text: 'Bu nedenle işbu davayı açma zorunluluğu doğmuştur.', kind: 'number', level: 0 },
            { text: 'Hizmet sözleşmesi', kind: 'bullet', level: 0 },
            { text: 'İhtarname ve tebliğ şerhi', kind: 'bullet', level: 0 },
            { text: 'Banka kayıtları ve faturalar', kind: 'bullet', level: 0 },
            { text: 'Tanık beyanları, bilirkişi incelemesi ve her türlü yasal delil', kind: 'bullet', level: 0 }
        ]
    };
}

/**
 * 2. İcra/itiraz dilekçesi: headings bold only through the Balk1 ("Başlık 1")
 * style, a two-level list (1. -> a)), a tracked deletion, a page break and
 * justified body text.
 */
function itirazDilekcesi() {
    const rec = recorder();
    const margins = { top: 1134, right: 1134, bottom: 1134, left: 1418 };
    const LIST = 4;
    const heading = (text, align) => p(rec.run(text), '<w:pStyle w:val="Balk1"/>' + (align ? jc(align) : ''));

    const body = [
        heading("ÖRNEK İCRA HUKUK MAHKEMESİ'NE", 'center'),
        empty(),
        p(rec.run('Esas No: 2026/0000') + rec.run('<w:tab/>', 'Takip Dosya No: 2026/0000'), jc('right')),
        empty(),
        party(rec, 'İTİRAZ EDEN', 'Ad Soyad (T.C. Kimlik No: 00000000000)', 'Örnek Mahallesi, Örnek Sokak No: 0, İlçe / İl'),
        party(rec, 'ALACAKLI', 'Örnek Finans A.Ş.'),
        party(rec, 'KONU', 'Ödeme emrine, borca ve faize itirazlarımızın sunulmasıdır.'),
        heading('İTİRAZ NEDENLERİ'),
        p(rec.run('Borca itiraz ediyoruz:'), list(LIST, 0)),
        p(rec.run('Takibe konu borcun ') + rec.del('tamamen ') + rec.run('ödendiği banka dekontuyla sabittir.'),
            list(LIST, 1) + jc('both')),
        p(rec.run('İşlemiş faiz ve uygulanan faiz oranı hukuka aykırıdır.'), list(LIST, 1) + jc('both')),
        p(rec.run('Yetkiye itiraz ediyoruz:'), list(LIST, 0)),
        p(rec.run('Borçlunun yerleşim yeri İlçe / İl olup takip yapılan icra dairesi yetkisizdir.'),
            list(LIST, 1) + jc('both')),
        // A page break the way Word writes Ctrl+Enter: a paragraph holding only the break
        p('<w:r><w:br w:type="page"/></w:r>'),
        heading('SONUÇ VE İSTEM', 'center'),
        p(rec.run('Açıklanan nedenlerle; ödeme emrine, borcun tamamına, faize, ferilerine ve yetkiye itiraz ' +
            'ettiğimizden takibin durdurulmasına, itirazlarımızın kabulüne ve alacaklı aleyhine yüzde yirmiden ' +
            'aşağı olmamak üzere tazminata hükmedilmesine karar verilmesini saygıyla arz ve talep ederim. ' +
            'ĞÜŞÖÇ harfleri büyük, ğüşöç harfleri küçük yazılmıştır.'), jc('both')),
        empty(),
        p(rec.run('İtiraz Eden'), jc('right')),
        p(rec.run('Ad Soyad'), jc('right'))
    ].join('');

    const numbering = abstractNum(0,
        level(0, 'decimal', '%1.', 720) + level(1, 'lowerLetter', '%2)', 1080) + level(2, 'lowerRoman', '%3.', 1440)) +
        num(LIST, 0);

    return {
        name: 'icra/itiraz dilekçesi',
        options: { styles: BASE_STYLES + HEADING_STYLE, numbering, sectPr: a4(margins) },
        body,
        ...recorded(rec),
        margins,
        headings: ["ÖRNEK İCRA HUKUK MAHKEMESİ'NE", 'İTİRAZ NEDENLERİ', 'SONUÇ VE İSTEM'],
        listItems: [
            { text: 'Borca itiraz ediyoruz:', kind: 'number', level: 0 },
            { text: 'Takibe konu borcun ', kind: 'number', level: 1 },
            { text: 'İşlemiş faiz ve uygulanan faiz oranı hukuka aykırıdır.', kind: 'number', level: 1 },
            { text: 'Yetkiye itiraz ediyoruz:', kind: 'number', level: 0 },
            { text: 'Borçlunun yerleşim yeri İlçe / İl olup takip yapılan icra dairesi yetkisizdir.', kind: 'number', level: 1 }
        ]
    };
}

/**
 * 3. Ekli tablo dilekçesi: underlined and italic runs, a receivables table with a
 * header row, a vertically merged cell, a line break in a cell and a horizontally
 * merged total row, then an EKLER list.
 */
function tabloDilekcesi() {
    const rec = recorder();
    const margins = { top: 1418, right: 1418, bottom: 1418, left: 1418 };
    const EKLER = 1;

    const cellPr = (width, extra = '') => `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>`;
    const cell = (width, content, { extra = '', align } = {}) =>
        `<w:tc>${cellPr(width, extra)}${p(content, align ? jc(align) : '')}</w:tc>`;
    const WIDTHS = [850, 2268, 3969, 1984];
    const row = (cells, trPr = '') => `<w:tr>${trPr}${cells.join('')}</w:tr>`;
    // description: the pieces of its run, recorded only when the row reaches that cell
    const dataRow = (no, type, description, amount, typeExtra = '') => row([
        cell(WIDTHS[0], rec.run(no), { align: 'center' }),
        // A continued vertical merge holds an empty paragraph, as Word writes it
        typeExtra.includes('w:vMerge/>')
            ? `<w:tc>${cellPr(WIDTHS[1], typeExtra)}<w:p/></w:tc>`
            : cell(WIDTHS[1], rec.run(type), { extra: typeExtra }),
        cell(WIDTHS[2], rec.run(...description)),
        cell(WIDTHS[3], rec.run(amount), { align: 'right' })
    ]);
    const shaded = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>';

    // Built when the body reaches it, so its text is recorded in document order
    const table = () => '<w:tbl><w:tblPr><w:tblStyle w:val="TabloKlavuzu"/><w:tblW w:w="0" w:type="auto"/>' +
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>' +
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
        `<w:tblGrid>${WIDTHS.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
        row(['Sıra', 'Alacak Türü', 'Açıklama', 'Tutar (TL)']
            .map((text, i) => cell(WIDTHS[i], rec.run(BOLD, text), { extra: shaded, align: 'center' })),
        '<w:trPr><w:tblHeader/></w:trPr>') +
        dataRow('1', 'Fatura alacağı',
            ['Örnek fatura no. 0001', '<w:br/>', '(00.00.2026 tarihli)'], '10.000,00', '<w:vMerge w:val="restart"/>') +
        dataRow('2', null, ['Örnek fatura no. 0002'], '5.250,50', '<w:vMerge/>') +
        dataRow('3', 'Gecikme faizi', ['Ödeme tarihine kadar işleyen yasal faiz'], '1.234,56') +
        dataRow('4', 'Masraflar', ['İhtarname ve tebligat giderleri'], '750,00') +
        row([
            `<w:tc>${cellPr(WIDTHS[0] + WIDTHS[1] + WIDTHS[2], '<w:gridSpan w:val="3"/>' + shaded)}` +
                `${p(rec.run(BOLD, 'TOPLAM'), jc('right'))}</w:tc>`,
            cell(WIDTHS[3], rec.run(BOLD, '17.235,06'), { extra: shaded, align: 'right' })
        ]) +
        '</w:tbl>';

    const body = [
        p(rec.run(BOLD, "ÖRNEK TÜKETİCİ MAHKEMESİ'NE"), jc('center')),
        p(rec.run('Esas No: 2026/0000'), jc('right')),
        empty(),
        party(rec, 'DAVACI', 'Ad Soyad (T.C. Kimlik No: 00000000000)'),
        party(rec, 'DAVALI', 'Örnek Ticaret A.Ş.'),
        empty(),
        p(rec.run('Davalıdan olan alacaklarımız aşağıdaki tabloda ') +
            rec.run({ rPr: '<w:u w:val="single"/>' }, 'kalem kalem') +
            rec.run(' gösterilmiştir; ') +
            rec.run({ rPr: '<w:i/><w:iCs/>' }, 'faiz hesabına ilişkin haklarımız saklı kalmak kaydıyla') +
            rec.run(' toplam tutarın davalıdan tahsilini talep ederiz.'), jc('both')),
        table(),
        empty(),
        p(rec.run(BOLD, 'EKLER'), ''),
        p(rec.run('Fatura örnekleri'), list(EKLER)),
        p(rec.run('Faiz hesap çizelgesi'), list(EKLER)),
        p(rec.run('Vekâletname örneği'), list(EKLER)),
        empty(),
        p(rec.run('Davacı'), jc('right')),
        p(rec.run('Ad Soyad'), jc('right'))
    ].join('');

    const numbering = abstractNum(0, [0, 1, 2].map(i => level(i, 'decimal', `%${i + 1}-`, 720 + i * 720)).join('')) +
        num(EKLER, 0);

    return {
        name: 'ekli tablo dilekçesi',
        options: { styles: BASE_STYLES, numbering, sectPr: a4(margins) },
        body,
        ...recorded(rec),
        margins,
        // Cells per row as UDF writes them: the continued vertical merge has no cell of its own
        tableShape: [4, 4, 3, 4, 4, 2],
        listItems: [
            { text: 'Fatura örnekleri', kind: 'number', level: 0 },
            { text: 'Faiz hesap çizelgesi', kind: 'number', level: 0 },
            { text: 'Vekâletname örneği', kind: 'number', level: 0 }
        ]
    };
}

const petitions = [davaDilekcesi(), itirazDilekcesi(), tabloDilekcesi()];
for (const petition of petitions) {
    petition.build = () => buildDocx(petition.body, petition.options);
}

module.exports = { petitions };
