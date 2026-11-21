const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const { buildDocx } = require('./helpers');
const { PAGE_URL, loadPage, fileOf, visible } = require('./helpers/page');

const ROOT = path.resolve(__dirname, '..');

const docxWith = text => buildDocx(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);

/**
 * Replace the app's parsers with ones whose results the test hands over by hand.
 * Each parse(file) call is recorded with resolve/reject for its promise, and
 * `real(file)` runs a real parser of its own.
 */
function stubParser({ window, run }) {
    const calls = [];
    let notify = null;
    window.app.createParser = () => ({
        parse: file => new Promise((resolve, reject) => {
            calls.push({ file, resolve, reject });
            if (notify) notify();
        })
    });
    return {
        calls,
        real: file => run('new DocxParser()').parse(file),
        // Resolves once the parser has been called `count` times in all
        async called(count) {
            while (calls.length < count) await new Promise(resolve => { notify = resolve; });
            notify = null;
            return calls[count - 1];
        }
    };
}

/** Record what the download button hands to downloadBlob instead of saving it */
function captureDownloads({ window, run }) {
    window.__downloads = [];
    run('downloadBlob = (blob, filename) => window.__downloads.push({ blob, filename })');
    return window.__downloads;
}

async function udfText(blob) {
    const udf = await JSZip.loadAsync(await blob.arrayBuffer());
    return udf.file('content.xml').async('string');
}

// Let every timer and promise the app has pending run
const settle = () => new Promise(resolve => setTimeout(resolve, 20));

/** The page shows the empty upload state and no result, error or notes */
function assertEmptyUploadState(document) {
    assert.ok(visible(document.getElementById('uploadZone')), 'upload zone visible');
    assert.ok(!visible(document.getElementById('fileInfo')), 'file row hidden');
    assert.ok(!visible(document.getElementById('progressContainer')), 'progress hidden');
    assert.ok(!visible(document.getElementById('downloadSection')), 'no download section');
    assert.ok(!visible(document.getElementById('conversionNotes')), 'no notes');
    assert.equal(document.querySelectorAll('#conversionNotesList > li').length, 0);
    assert.ok(!visible(document.getElementById('errorMessage')), 'no error');
    const convertBtn = document.getElementById('convertBtn');
    assert.ok(visible(convertBtn), 'convert button visible');
    assert.equal(convertBtn.disabled, true, 'convert button disabled');
    assert.equal(document.getElementById('progressFill').style.width, '0%', 'bar reset');
    assert.equal(document.getElementById('progressBar').getAttribute('aria-valuenow'), '0');
}

test('app.js waits on no artificial delays, and utils.js no longer has sleep()', () => {
    const app = fs.readFileSync(path.join(ROOT, 'js/app.js'), 'utf8');
    assert.doesNotMatch(app, /\bsleep\(/);
    const utils = fs.readFileSync(path.join(ROOT, 'js/utils.js'), 'utf8');
    assert.doesNotMatch(utils, /function sleep\b/);
});

test('removing the file mid-conversion cancels it: a result that arrives later is dropped', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);
    const docx = await docxWith('Dilekçe A');

    window.app.processFile(fileOf('a.docx', docx));
    const conversion = window.app.handleConvert();
    const call = await parser.called(1);
    assert.ok(visible(document.getElementById('progressContainer')));
    assert.ok(!visible(document.getElementById('convertBtn')));

    // KALDIR stays usable during the conversion
    const remove = document.getElementById('removeFile');
    assert.ok(visible(document.getElementById('fileInfo')));
    assert.equal(remove.disabled, false);
    remove.click();
    assertEmptyUploadState(document);

    call.resolve(await parser.real(call.file));
    await conversion;
    await settle();

    assertEmptyUploadState(document);
    assert.equal(window.app.convertedBlob, null);
    assert.equal(window.app.selectedFile, null);
    assert.equal(window.app.outputFilename, 'a.udf', 'left as the file selection set it');
    assert.deepEqual(page.errors, []);
});

test('a parse that fails after its file was removed shows no error', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);

    window.app.processFile(fileOf('a.docx', await docxWith('Dilekçe A')));
    const conversion = window.app.handleConvert();
    const call = await parser.called(1);
    document.getElementById('removeFile').click();

    const failure = new Error('Invalid or corrupted DOCX file');
    failure.code = 'corrupt';
    call.reject(failure);
    await conversion;
    await settle();

    assertEmptyUploadState(document);
    assert.equal(document.getElementById('errorText').textContent, page.strings.tr['error.default'],
        'the error text was never set');
    assert.deepEqual(page.errors, [], 'nothing logged for a conversion nobody is waiting for');
});

test('removing file A mid-conversion and choosing B: A never shows, and B converts to B.udf with B\'s text', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);
    const downloads = captureDownloads(page);

    window.app.processFile(fileOf('a.docx', await docxWith('Dilekçe A')));
    const conversionA = window.app.handleConvert();
    const callA = await parser.called(1);

    document.getElementById('removeFile').click();
    window.app.processFile(fileOf('b.docx', await docxWith('Dilekçe B')));
    assert.equal(document.getElementById('fileName').textContent, 'b.docx');

    // A finishes while B waits to be converted: nothing of A appears
    callA.resolve(await parser.real(callA.file));
    await conversionA;
    await settle();
    assert.ok(!visible(document.getElementById('downloadSection')));
    assert.ok(!visible(document.getElementById('progressContainer')));
    assert.ok(visible(document.getElementById('fileInfo')));
    assert.equal(document.getElementById('convertBtn').disabled, false);
    assert.equal(window.app.convertedBlob, null);

    const conversionB = window.app.handleConvert();
    const callB = await parser.called(2);
    assert.equal(callB.file.name, 'b.docx');
    callB.resolve(await parser.real(callB.file));
    await conversionB;

    assert.ok(visible(document.getElementById('downloadSection')));
    document.getElementById('downloadBtn').click();
    assert.equal(downloads.length, 1);
    assert.equal(downloads[0].filename, 'b.udf');
    const xml = await udfText(downloads[0].blob);
    assert.match(xml, /Dilekçe B/);
    assert.doesNotMatch(xml, /Dilekçe A/);
});

test('a result for a removed file that arrives while the next file converts is dropped too', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);
    const downloads = captureDownloads(page);

    window.app.processFile(fileOf('a.docx', await docxWith('Dilekçe A')));
    const conversionA = window.app.handleConvert();
    const callA = await parser.called(1);
    document.getElementById('removeFile').click();

    window.app.processFile(fileOf('b.docx', await docxWith('Dilekçe B')));
    const conversionB = window.app.handleConvert();
    const callB = await parser.called(2);

    callA.resolve(await parser.real(callA.file));
    await conversionA;
    await settle();
    assert.ok(!visible(document.getElementById('downloadSection')), 'A is not shown');
    assert.ok(visible(document.getElementById('progressContainer')), 'B is still converting');
    assert.equal(window.app.convertedBlob, null);

    callB.resolve(await parser.real(callB.file));
    await conversionB;
    assert.ok(visible(document.getElementById('downloadSection')));
    document.getElementById('downloadBtn').click();
    assert.equal(downloads[0].filename, 'b.udf');
    assert.match(await udfText(downloads[0].blob), /Dilekçe B/);
});

test('each conversion gets its own parser and generator, so a discarded run cannot touch the next one', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parsers = [];
    const generators = [];
    const { createParser, createGenerator } = window.app;
    let parseStarted;
    const aParsing = new Promise(resolve => { parseStarted = resolve; });
    window.app.createParser = () => {
        const parser = createParser.call(window.app);
        const parse = parser.parse.bind(parser);
        parser.parse = file => { const result = parse(file); parseStarted(); return result; };
        parsers.push(parser);
        return parser;
    };
    window.app.createGenerator = () => { const generator = createGenerator.call(window.app); generators.push(generator); return generator; };
    const downloads = captureDownloads(page);

    // A's real parse is still running when B's starts: both run at once
    const bigA = await buildDocx(Array.from({ length: 300 }, (_, i) => `<w:p><w:r><w:t>Dilekçe A ${i}</w:t></w:r></w:p>`).join(''));
    window.app.processFile(fileOf('a.docx', bigA));
    const docxB = await docxWith('Dilekçe B');
    const conversionA = window.app.handleConvert();
    await aParsing;
    document.getElementById('removeFile').click();
    window.app.processFile(fileOf('b.docx', docxB));
    const conversionB = window.app.handleConvert();
    await Promise.all([conversionA, conversionB]);

    assert.equal(new Set(parsers).size, 2);
    assert.equal(new Set(generators).size, 2);
    assert.ok(visible(document.getElementById('downloadSection')));
    document.getElementById('downloadBtn').click();
    assert.equal(downloads.length, 1);
    assert.equal(downloads[0].filename, 'b.udf');
    const xml = await udfText(downloads[0].blob);
    assert.match(xml, /Dilekçe B/);
    assert.doesNotMatch(xml, /Dilekçe A/);
});

test('"convert another" also drops a conversion still running', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);

    window.app.processFile(fileOf('a.docx', await docxWith('Dilekçe A')));
    const conversion = window.app.handleConvert();
    const call = await parser.called(1);
    window.app.handleConvertAnother();

    call.resolve(await parser.real(call.file));
    await conversion;
    await settle();
    assertEmptyUploadState(document);
    assert.equal(window.app.convertedBlob, null);
});

test('the progress bar and stage text are exposed to assistive technology, in both languages', async () => {
    const { document, strings } = await loadPage({ app: true });
    const bar = document.getElementById('progressBar');
    assert.equal(bar.getAttribute('role'), 'progressbar');
    assert.equal(bar.getAttribute('aria-valuemin'), '0');
    assert.equal(bar.getAttribute('aria-valuemax'), '100');
    assert.equal(bar.getAttribute('aria-valuenow'), '0');
    assert.ok(bar.contains(document.getElementById('progressFill')));
    assert.equal(document.getElementById('progressText').getAttribute('role'), 'status');

    assert.ok(strings.tr['progress.label'] && strings.en['progress.label']);
    assert.notEqual(strings.tr['progress.label'], strings.en['progress.label']);
    assert.equal(bar.getAttribute('aria-label'), strings.tr['progress.label']);
    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(bar.getAttribute('aria-label'), strings.en['progress.label']);
});

test('the status region is shown empty and gets the first stage a frame later, so it is announced', async () => {
    const page = await loadPage({ app: true });
    const { window, document, strings } = page;
    const parser = stubParser(page);
    const text = document.getElementById('progressText');

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni')));
    const conversion = window.app.handleConvert();
    assert.ok(visible(document.getElementById('progressContainer')));
    assert.equal(text.textContent, '', 'empty when it appears');

    // A language change in that frame leaves it empty too
    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(text.textContent, '');

    const call = await parser.called(1);
    assert.equal(text.textContent, strings.en['progress.parsing'], 'filled before the parse starts');
    call.resolve(await parser.real(call.file));
    await conversion;
    assert.ok(visible(document.getElementById('downloadSection')));
});

test('each real step shows its stage, and aria-valuenow follows the bar', async () => {
    const page = await loadPage({ url: `${PAGE_URL}?lang=en`, app: true });
    const { window, document, strings } = page;
    const bar = document.getElementById('progressBar');
    const fill = document.getElementById('progressFill');
    const text = document.getElementById('progressText');
    const seen = [];
    const snapshot = step => seen.push({ step, text: text.textContent, now: bar.getAttribute('aria-valuenow'), width: fill.style.width });

    const { createParser, createGenerator } = window.app;
    window.app.createParser = () => {
        const parser = createParser.call(window.app);
        const parse = parser.parse.bind(parser);
        parser.parse = file => { snapshot('parse'); return parse(file); };
        return parser;
    };
    window.app.createGenerator = () => {
        const generator = createGenerator.call(window.app);
        const generate = generator.generate.bind(generator);
        generator.generate = doc => { snapshot('generate'); return generate(doc); };
        return generator;
    };

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni')));
    await window.app.handleConvert();
    snapshot('done');

    assert.deepEqual(seen, [
        { step: 'parse', text: strings.en['progress.parsing'], now: '33', width: '33%' },
        { step: 'generate', text: strings.en['progress.converting'], now: '67', width: '67%' },
        { step: 'done', text: strings.en['progress.complete'], now: '100', width: '100%' }
    ]);
    assert.ok(visible(document.getElementById('downloadSection')));

    // "Convert another" puts the bar back at zero
    document.getElementById('convertAnother').click();
    assert.equal(bar.getAttribute('aria-valuenow'), '0');
    assert.equal(fill.style.width, '0%');
});

test('a successful conversion moves focus to the download button', async () => {
    const { window, document } = await loadPage({ app: true });
    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni')));
    document.getElementById('convertBtn').focus();
    await window.app.handleConvert();

    assert.ok(visible(document.getElementById('downloadSection')));
    assert.equal(document.activeElement, document.getElementById('downloadBtn'));
});

test('a failed conversion shows the alert and leaves focus on the convert button', async () => {
    const { window, document } = await loadPage({ app: true });
    window.app.processFile(fileOf('bozuk.docx', 'this is not a zip archive'));
    await window.app.handleConvert();

    assert.ok(visible(document.getElementById('errorMessage')));
    assert.equal(document.getElementById('errorMessage').getAttribute('role'), 'alert');
    assert.ok(visible(document.getElementById('convertBtn')));
    assert.equal(document.activeElement, document.getElementById('convertBtn'));
});

// Source lines of the result, as shown
const sourceLines = document => ['resultSourceFile', 'resultConverted', 'resultOutput']
    .map(id => document.getElementById(id).textContent);

const MOMENT_OPTIONS = {
    dateTime: { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' },
    time: { hour: '2-digit', minute: '2-digit', second: '2-digit' }
};
const moment = (style, ms, lang) => new Intl.DateTimeFormat(lang, MOMENT_OPTIONS[style]).format(ms);

// Midday UTC, so these are 21 November in any time zone within ±11 hours
const REVISION_1 = Date.UTC(2026, 10, 21, 12, 3, 9);
const REVISION_2 = Date.UTC(2026, 10, 21, 12, 41, 37);
const CONVERTED_AT = Date.UTC(2026, 10, 21, 12, 45, 2);

/** A stub-parser result for `call`: the real parse of its file, with the given findings */
async function parsedWith(parser, call, findings) {
    return { ...(await parser.real(call.file)), findings };
}

/** Choose `file` in the replace picker, the way the browser does after the button opens it */
function pickReplacement(window, document, file) {
    const input = document.getElementById('replaceInput');
    Object.defineProperty(input, 'files', { configurable: true, value: file ? [file] : [] });
    input.dispatchEvent(new window.Event('change'));
}

test('the result names its source file, size, last-modified time, conversion time and .udf name, in TR and then EN', async () => {
    const page = await loadPage({ app: true });
    const { window, document, strings, run } = page;
    run(`Date.now = () => ${CONVERTED_AT}`);
    const docx = await docxWith('Dilekçe metni');

    window.app.processFile(fileOf('dilekce.docx', docx, REVISION_1));
    await window.app.handleConvert();
    assert.ok(visible(document.getElementById('downloadSection')));

    const expected = lang => [
        strings[lang]['result.source']
            .replace('{name}', 'dilekce.docx')
            .replace('{size}', run(`formatFileSize(${docx.length})`))
            .replace('{modified}', moment('dateTime', REVISION_1, lang)),
        strings[lang]['result.converted'].replace('{time}', moment('time', CONVERTED_AT, lang)),
        strings[lang]['result.output'].replace('{output}', 'dilekce.udf')
    ];
    assert.deepEqual(sourceLines(document), expected('tr'));
    assert.match(sourceLines(document)[0], /^Kaynak: dilekce\.docx \(.+ KB, değiştirilme 21\.11\.2026 \d\d:\d\d:09\)$/);
    assert.equal(sourceLines(document)[2], 'İndirilecek: dilekce.udf');

    // The language toggle writes the lines, and their times, again while the result is shown
    document.querySelector('[data-set-lang="en"]').click();
    assert.deepEqual(sourceLines(document), expected('en'));
    assert.match(sourceLines(document)[0], /^Source: dilekce\.docx \(.+ KB, modified 11\/21\/2026, \d\d:\d\d:09 [AP]M\)$/);
    document.querySelector('[data-set-lang="tr"]').click();
    assert.deepEqual(sourceLines(document), expected('tr'));

    // Plain text read with the result, in no live region of its own
    const source = document.getElementById('resultSource');
    assert.ok(document.getElementById('downloadSection').contains(source));
    for (let el = source; el; el = el.parentElement) {
        assert.ok(!el.hasAttribute('role') && !el.hasAttribute('aria-live'), `${el.id || el.localName} is no live region`);
    }
});

test('the values come from the file the run converted, not the file selected when the result shows', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni'), REVISION_1));
    const conversion = window.app.handleConvert();
    const call = await parser.called(1);
    // Mid-run, the app's selected file is swapped without a new run (no UI does this; it checks the capture)
    window.app.selectedFile = fileOf('baska.docx', await docxWith('Başka'), REVISION_2);
    call.resolve(await parser.real(call.file));
    await conversion;
    assert.match(sourceLines(document)[0], /dilekce\.docx/);
    assert.match(sourceLines(document)[0], new RegExp(moment('dateTime', REVISION_1, 'tr').replace(/\./g, '\\.')));
    assert.equal(sourceLines(document)[2], 'İndirilecek: dilekce.udf');
});

test('a file without a last-modified time gets a source line without one', async () => {
    const { window, document, strings } = await loadPage({ app: true });
    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni')));
    await window.app.handleConvert();
    assert.match(sourceLines(document)[0], /^Kaynak: dilekce\.docx \([^,]+\)$/);
    assert.equal(document.getElementById('resultSourceFile').dataset.i18n, 'result.sourceUndated');
    assert.ok(strings.en['result.sourceUndated']);
});

test('a file name with markup in it is shown as literal text', async () => {
    const { window, document } = await loadPage({ app: true });
    const name = '<img src=x onerror=alert(1)>.docx';
    window.app.processFile(fileOf(name, await docxWith('Dilekçe metni'), REVISION_1));
    await window.app.handleConvert();

    assert.ok(visible(document.getElementById('downloadSection')));
    assert.ok(sourceLines(document)[0].includes(name));
    assert.equal(sourceLines(document)[2], 'İndirilecek: <img src=x onerror=alert(1)>.udf');
    assert.equal(document.querySelectorAll('#downloadSection img').length, 0);
    document.querySelector('[data-set-lang="en"]').click();
    assert.ok(sourceLines(document)[0].includes(name));
    assert.equal(document.querySelectorAll('#downloadSection img').length, 0);
});

test('"replace source" with a same-name revision drops the old download and notes at once, then shows only the new one', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);
    const downloads = captureDownloads(page);

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe eski sürüm'), REVISION_1));
    const first = window.app.handleConvert();
    const call1 = await parser.called(1);
    call1.resolve(await parsedWith(parser, call1, [{ code: 'comments', count: 2, examples: [] }]));
    await first;
    assert.ok(visible(document.getElementById('conversionNotes')));
    const oldSource = sourceLines(document)[0];

    // The revision: same name, new text, later last-modified time, chosen through the picker
    pickReplacement(window, document, fileOf('dilekce.docx', await docxWith('Dilekçe yeni sürüm'), REVISION_2));

    // Before the new conversion resolves, nothing of the old result is left
    const call2 = await parser.called(2);
    assert.equal(call2.file.lastModified, REVISION_2);
    assert.ok(!visible(document.getElementById('downloadSection')), 'old download gone');
    assert.equal(window.app.convertedBlob, null);
    assert.ok(!visible(document.getElementById('conversionNotes')), 'old notes gone');
    assert.equal(document.querySelectorAll('#conversionNotesList > li').length, 0);
    assert.deepEqual(sourceLines(document), ['', '', ''], 'old source line gone');
    assert.ok(visible(document.getElementById('progressContainer')), 'the new conversion runs with its progress');
    assert.ok(!visible(document.getElementById('convertBtn')));
    assert.equal(document.getElementById('fileName').textContent, 'dilekce.docx');

    call2.resolve(await parsedWith(parser, call2, []));
    await settle();
    assert.ok(visible(document.getElementById('downloadSection')));
    assert.equal(document.activeElement, document.getElementById('downloadBtn'));
    assert.ok(!visible(document.getElementById('conversionNotes')), 'the new result has no notes');
    assert.notEqual(sourceLines(document)[0], oldSource);
    assert.ok(sourceLines(document)[0].includes(moment('dateTime', REVISION_2, 'tr')));
    assert.ok(!sourceLines(document)[0].includes(moment('dateTime', REVISION_1, 'tr')));

    document.getElementById('downloadBtn').click();
    assert.equal(downloads.length, 1);
    assert.equal(downloads[0].filename, 'dilekce.udf');
    const xml = await udfText(downloads[0].blob);
    assert.match(xml, /Dilekçe yeni sürüm/);
    assert.doesNotMatch(xml, /eski sürüm/);
    assert.deepEqual(page.errors, []);
});

test('the replace button empties its picker before opening it, so the same file name counts as a new choice', async () => {
    const { window, document } = await loadPage({ app: true });
    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni'), REVISION_1));
    await window.app.handleConvert();

    const input = document.getElementById('replaceInput');
    let opened = 0;
    input.click = () => { opened++; };
    let cleared = false;
    Object.defineProperty(input, 'value', { configurable: true, get: () => '', set: value => { cleared = value === ''; } });
    document.getElementById('replaceSource').click();
    assert.equal(opened, 1);
    assert.ok(cleared);
    assert.equal(input.getAttribute('type'), 'file');
    assert.equal(input.getAttribute('accept'), '.docx');
});

test('replacing with a file whose conversion fails shows the error and none of the old result', async () => {
    const page = await loadPage({ app: true });
    const { window, document, strings } = page;

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni'), REVISION_1));
    await window.app.handleConvert();
    assert.ok(visible(document.getElementById('downloadSection')));

    // word/document.xml cut off mid-run
    const zip = await JSZip.loadAsync(await docxWith('Dilekçe metni'));
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Dilek');
    const broken = await zip.generateAsync({ type: 'nodebuffer' });
    await window.app.replaceSource(fileOf('dilekce.docx', broken, REVISION_2));

    assert.ok(visible(document.getElementById('errorMessage')));
    assert.equal(document.getElementById('errorText').textContent, strings.tr['error.corrupt']);
    assert.ok(!visible(document.getElementById('downloadSection')));
    assert.equal(window.app.convertedBlob, null);
    assert.deepEqual(sourceLines(document), ['', '', '']);
    assert.ok(visible(document.getElementById('convertBtn')));
    assert.equal(document.getElementById('convertBtn').disabled, false);
    assert.ok(visible(document.getElementById('fileInfo')), 'the new file stays selected');
    assert.equal(window.app.selectedFile.lastModified, REVISION_2);
    assert.equal(document.activeElement, document.getElementById('convertBtn'));
});

test('cancelling the replace picker leaves the result alone; a non-.docx replacement keeps it and says why', async () => {
    const page = await loadPage({ app: true });
    const { window, document, strings } = page;
    const downloads = captureDownloads(page);

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni'), REVISION_1));
    await window.app.handleConvert();
    const blob = window.app.convertedBlob;
    const lines = sourceLines(document);

    // The picker opens and is closed with no file: no change event, or one with no files
    document.getElementById('replaceSource').click();
    pickReplacement(window, document, null);
    assert.ok(visible(document.getElementById('downloadSection')));
    assert.equal(window.app.convertedBlob, blob);
    assert.deepEqual(sourceLines(document), lines);
    assert.ok(!visible(document.getElementById('errorMessage')));

    pickReplacement(window, document, fileOf('notlar.txt', 'düz metin', REVISION_2));
    await settle();
    assert.ok(visible(document.getElementById('errorMessage')));
    assert.equal(document.getElementById('errorText').textContent, strings.tr['error.invalidType']);
    assert.ok(visible(document.getElementById('downloadSection')));
    assert.equal(window.app.convertedBlob, blob);
    assert.deepEqual(sourceLines(document), lines);
    document.getElementById('downloadBtn').click();
    assert.equal(downloads[0].filename, 'dilekce.udf');
    assert.match(await udfText(downloads[0].blob), /Dilekçe metni/);
});

test('replacing again while a replacement still converts drops the earlier one', async () => {
    const page = await loadPage({ app: true });
    const { window, document } = page;
    const parser = stubParser(page);
    const downloads = captureDownloads(page);

    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe A'), REVISION_1));
    const first = window.app.handleConvert();
    const callA = await parser.called(1);
    callA.resolve(await parser.real(callA.file));
    await first;

    window.app.replaceSource(fileOf('dilekce.docx', await docxWith('Dilekçe B'), REVISION_2));
    const callB = await parser.called(2);
    window.app.replaceSource(fileOf('dilekce.docx', await docxWith('Dilekçe C'), CONVERTED_AT));
    const callC = await parser.called(3);

    // B finishes first: it is stale and shows nothing
    callB.resolve(await parsedWith(parser, callB, [{ code: 'comments', count: 1, examples: [] }]));
    await settle();
    assert.ok(!visible(document.getElementById('downloadSection')));
    assert.equal(window.app.convertedBlob, null);
    assert.deepEqual(sourceLines(document), ['', '', '']);
    assert.ok(visible(document.getElementById('progressContainer')), 'C still converts');

    callC.resolve(await parsedWith(parser, callC, []));
    await settle();
    assert.ok(visible(document.getElementById('downloadSection')));
    assert.ok(!visible(document.getElementById('conversionNotes')), 'B\'s notes never show');
    assert.ok(sourceLines(document)[0].includes(moment('dateTime', CONVERTED_AT, 'tr')));
    document.getElementById('downloadBtn').click();
    assert.equal(downloads.length, 1);
    const xml = await udfText(downloads[0].blob);
    assert.match(xml, /Dilekçe C/);
    assert.doesNotMatch(xml, /Dilekçe [AB]/);

    // A failing stale run reports nothing either
    window.app.replaceSource(fileOf('dilekce.docx', await docxWith('Dilekçe D'), REVISION_1));
    const callD = await parser.called(4);
    window.app.replaceSource(fileOf('dilekce.docx', await docxWith('Dilekçe E'), REVISION_2));
    const callE = await parser.called(5);
    callD.reject(Object.assign(new Error('Invalid or corrupted DOCX file'), { code: 'corrupt' }));
    await settle();
    assert.ok(!visible(document.getElementById('errorMessage')));
    callE.resolve(await parser.real(callE.file));
    await settle();
    assert.ok(visible(document.getElementById('downloadSection')));
    assert.deepEqual(page.errors, []);
});

test('"convert another" clears the source line along with the rest of the result', async () => {
    const { window, document } = await loadPage({ app: true });
    window.app.processFile(fileOf('dilekce.docx', await docxWith('Dilekçe metni'), REVISION_1));
    await window.app.handleConvert();
    assert.notEqual(sourceLines(document)[0], '');

    document.getElementById('convertAnother').click();
    assertEmptyUploadState(document);
    assert.deepEqual(sourceLines(document), ['', '', '']);
    for (const id of ['resultSourceFile', 'resultConverted', 'resultOutput']) {
        assert.equal(document.getElementById(id).dataset.i18n, undefined, `${id} is not re-filled on a language change`);
    }
    document.querySelector('[data-set-lang="en"]').click();
    assert.deepEqual(sourceLines(document), ['', '', '']);
});

test('the replace button is a keyboard-reachable button named in TR and EN', async () => {
    const { document, strings } = await loadPage({ app: true });
    const button = document.getElementById('replaceSource');
    assert.equal(button.localName, 'button');
    assert.equal(button.getAttribute('type'), 'button');
    assert.ok(!button.hasAttribute('tabindex') && !button.disabled && !button.hasAttribute('aria-hidden'));
    assert.ok(document.getElementById('downloadSection').contains(button));
    assert.ok(strings.tr['download.replace'] && strings.en['download.replace']);
    assert.notEqual(strings.tr['download.replace'], strings.en['download.replace']);
    assert.equal(button.textContent, strings.tr['download.replace']);
    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(button.textContent, strings.en['download.replace']);
    assert.equal(document.getElementById('replaceInput').getAttribute('aria-label'), strings.en['download.replaceInputLabel']);
});
