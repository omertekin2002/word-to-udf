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
    window.app.createParser = () => { const parser = createParser.call(window.app); parsers.push(parser); return parser; };
    window.app.createGenerator = () => { const generator = createGenerator.call(window.app); generators.push(generator); return generator; };
    const downloads = captureDownloads(page);

    // A's real parse is still running when B's starts: both run at once
    const bigA = await buildDocx(Array.from({ length: 300 }, (_, i) => `<w:p><w:r><w:t>Dilekçe A ${i}</w:t></w:r></w:p>`).join(''));
    window.app.processFile(fileOf('a.docx', bigA));
    const conversionA = window.app.handleConvert();
    document.getElementById('removeFile').click();
    window.app.processFile(fileOf('b.docx', await docxWith('Dilekçe B')));
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
