/**
 * Browser smoke test: serve the site from 127.0.0.1, open index.html in headless
 * Chromium and drive a whole conversion through the real page, its Content-Security-Policy,
 * the vendored JSZip and the browser's own DOMParser.
 *
 * Needs a Playwright Chromium (`bunx playwright install chromium`); run with `bun run test:browser`.
 */

const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const http = require('node:http');
const path = require('node:path');
const JSZip = require('jszip');
const { JSDOM } = require('jsdom');
const { chromium } = require('playwright');
const { buildDocx, declareSize } = require('../helpers');
const { petitions } = require('../fixtures/petitions');

const ROOT = path.resolve(__dirname, '..', '..');
const { DOMParser } = new JSDOM('').window;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
// Long enough for the page's progress delays, short enough that a broken step fails fast
const STEP_TIMEOUT = 15000;

const CONTENT_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.png': 'image/png'
};

/**
 * Serve the repo root read-only on 127.0.0.1 at a random port.
 * @returns {Promise<http.Server>}
 */
function startServer() {
    const server = http.createServer(async (req, res) => {
        try {
            const pathname = decodeURIComponent(new URL(req.url, 'http://127.0.0.1').pathname);
            const file = path.resolve(ROOT, '.' + (pathname.endsWith('/') ? pathname + 'index.html' : pathname));
            if ((req.method !== 'GET' && req.method !== 'HEAD') || !file.startsWith(ROOT + path.sep)) {
                throw new Error('not served');
            }
            const body = await fs.readFile(file);
            res.writeHead(200, { 'Content-Type': CONTENT_TYPES[path.extname(file)] || 'application/octet-stream' });
            res.end(req.method === 'HEAD' ? undefined : body);
        } catch {
            res.writeHead(404);
            res.end();
        }
    });
    return new Promise((resolve, reject) => {
        server.once('error', reject);
        server.listen(0, '127.0.0.1', () => resolve(server));
    });
}

let server;
let browser;
let origin;

before(async () => {
    server = await startServer();
    assert.equal(server.address().address, '127.0.0.1', 'the test server listens on loopback only');
    origin = `http://127.0.0.1:${server.address().port}`;
    browser = await chromium.launch();
});

// Runs even when a test fails, so no browser or server outlives the run
after(async () => {
    await browser?.close();
    if (server) await new Promise(resolve => server.close(resolve));
});

/**
 * Open the page in a fresh context, recording every console error, page error,
 * CSP violation, failed or off-origin request. The context closes when the test ends.
 * @param {import('node:test').TestContext} t
 * @param {string} [query] - e.g. '?lang=en'
 */
async function openPage(t, query = '') {
    const context = await browser.newContext({ acceptDownloads: true });
    t.after(() => context.close());
    const page = await context.newPage();
    page.setDefaultTimeout(STEP_TIMEOUT);

    const problems = { consoleErrors: [], pageErrors: [], cspViolations: [], failedRequests: [], foreignRequests: [] };
    page.on('console', message => {
        if (message.type() === 'error') problems.consoleErrors.push(message.text());
    });
    page.on('pageerror', error => problems.pageErrors.push(error.message));
    page.on('request', request => {
        // blob: URLs carry the origin that created them
        if (new URL(request.url()).origin !== origin) problems.foreignRequests.push(request.url());
    });
    page.on('requestfailed', request => problems.failedRequests.push(`${request.url()} ${request.failure()?.errorText}`));
    page.on('response', response => {
        if (response.status() >= 400) problems.failedRequests.push(`${response.url()} ${response.status()}`);
    });
    // Injected through the DevTools protocol, so the page's CSP doesn't block it
    await page.exposeFunction('reportCspViolation', violation => problems.cspViolations.push(violation));
    await page.addInitScript(() => {
        document.addEventListener('securitypolicyviolation', event => {
            window.reportCspViolation(`${event.violatedDirective} blocked ${event.blockedURI || 'inline'} ` +
                `(${event.sourceFile || ''}:${event.lineNumber})`);
        });
    });

    await page.goto(`${origin}/index.html${query}`, { waitUntil: 'load' });
    return { page, problems };
}

/**
 * Nothing went wrong on the page. `allowConsoleErrors` matches console errors a test expects.
 */
function assertClean(problems, { allowConsoleErrors = [] } = {}) {
    const consoleErrors = problems.consoleErrors
        .filter(text => !allowConsoleErrors.some(pattern => pattern.test(text)));
    assert.deepEqual({ ...problems, consoleErrors }, {
        consoleErrors: [], pageErrors: [], cspViolations: [], failedRequests: [], foreignRequests: []
    });
}

/** Choose a .docx through the file input, the way the browse dialog does */
async function chooseFile(page, name, buffer) {
    await page.locator('#fileInput').setInputFiles({ name, mimeType: DOCX_TYPE, buffer });
}

/** Choose a file, convert it and wait for the download section */
async function convertFile(page, name, buffer) {
    await chooseFile(page, name, buffer);
    await page.locator('#convertBtn:enabled').click();
    await page.locator('#downloadSection').waitFor({ state: 'visible' });
}

/** A small, valid .docx with no conversion findings */
const plainDocx = () => buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>');

test('the page loads with its scripts, styles and CSP, and nothing fails or leaves the origin', async t => {
    const { page, problems } = await openPage(t);

    // Every page script ran: the vendored JSZip, the converter classes and the app wiring
    assert.deepEqual(await page.evaluate(() => ({
        jszip: typeof JSZip,
        parser: typeof DocxParser,
        generator: typeof UdfGenerator,
        app: Boolean(window.app)
    })), { jszip: 'function', parser: 'function', generator: 'function', app: true });

    // The stylesheet applied: hidden sections are not displayed
    assert.equal(await page.locator('#downloadSection').evaluate(el => getComputedStyle(el).display), 'none');
    assert.equal(await page.locator('#convertBtn').isDisabled(), true);
    assertClean(problems);
});

test('a whole petition converts and downloads as a .udf holding its text', async t => {
    const { page, problems } = await openPage(t);
    const petition = petitions[0];

    await chooseFile(page, 'dava-dilekcesi.docx', await petition.build());
    assert.equal(await page.locator('#fileName').textContent(), 'dava-dilekcesi.docx');
    assert.equal(await page.locator('#fileInfo').isVisible(), true);
    assert.equal(await page.locator('#convertBtn').isEnabled(), true);

    await page.locator('#convertBtn').click();
    await page.locator('#downloadSection').waitFor({ state: 'visible' });
    assert.equal(await page.locator('#errorMessage').isVisible(), false);

    const [download] = await Promise.all([
        page.waitForEvent('download'),
        page.locator('#downloadBtn').click()
    ]);
    assert.equal(download.suggestedFilename(), 'dava-dilekcesi.udf');

    const udf = await JSZip.loadAsync(await fs.readFile(await download.path()));
    const xml = await udf.file('content.xml')?.async('string');
    assert.ok(xml, 'the .udf holds a content.xml');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    assert.equal(doc.getElementsByTagName('parsererror').length, 0, 'content.xml must be well-formed');
    const content = doc.querySelector('template > content').textContent;
    for (const text of [...petition.texts, ...petition.footnotes]) {
        assert.ok(content.includes(text), `content.xml is missing ${JSON.stringify(text)}`);
    }
    assertClean(problems);
});

test('a header is written into the .udf; a page number in the footer shows one conversion note; a file with no findings shows none', async t => {
    const { page, problems } = await openPage(t);

    const pageField = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const withHeader = await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>', {
        sectPr: '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>',
        parts: [
            { name: 'header1.xml', root: 'hdr', type: 'header', id: 'rIdHeader1',
                inner: '<w:p><w:r><w:t>Ad Soyad Hukuk Bürosu</w:t></w:r></w:p>' },
            { name: 'footer1.xml', root: 'ftr', type: 'footer', id: 'rIdFooter1',
                inner: `<w:p><w:r><w:t xml:space="preserve">Dosya No: 2026/0000 — Sayfa </w:t></w:r>${pageField}</w:p>` }
        ]
    });
    await convertFile(page, 'ust-bilgi.docx', withHeader);
    await page.locator('#conversionNotes').waitFor({ state: 'visible' });
    const notes = page.locator('#conversionNotesList > li.conversion-note');
    assert.equal(await notes.count(), 1);
    assert.equal(await notes.first().getAttribute('data-code'), 'headersFooters');
    assert.equal(await notes.first().locator('.conversion-note-text').textContent(),
        await page.evaluate(() => t('notes.headersFooters')));

    // Converted by Chromium's own XML parser: header first, footer last, the field's result left out
    const xml = await page.evaluate(async () =>
        (await JSZip.loadAsync(await window.app.convertedBlob.arrayBuffer())).file('content.xml').async('string'));
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    assert.equal(doc.getElementsByTagName('parsererror').length, 0, 'content.xml must be well-formed');
    assert.equal(doc.querySelector('template > content').textContent,
        'Ad Soyad Hukuk Bürosu' + 'Dilekçe metni' + 'Dosya No: 2026/0000 — Sayfa ');
    assert.deepEqual([...doc.querySelector('elements').children].map(el => el.localName), ['header', 'paragraph', 'footer']);

    await page.locator('#convertAnother').click();
    await convertFile(page, 'duz.docx', await plainDocx());
    assert.equal(await page.locator('#conversionNotes').isVisible(), false);
    assert.equal(await page.locator('#conversionNotesList > li').count(), 0);
    assertClean(problems);
});

test('a .docx with a malformed word/document.xml is refused as damaged', async t => {
    const { page, problems } = await openPage(t);

    // Truncated mid-run: the browser's DOMParser decides whether this is well-formed
    const zip = await JSZip.loadAsync(await plainDocx());
    zip.file('word/document.xml',
        `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Dilek`);
    await chooseFile(page, 'bozuk.docx', await zip.generateAsync({ type: 'nodebuffer' }));
    await page.locator('#convertBtn').click();

    await page.locator('#errorMessage').waitFor({ state: 'visible' });
    const expected = await page.evaluate(() => I18N_STRINGS.tr['error.corrupt']);
    assert.ok(expected, 'the dictionary has an error.corrupt message');
    assert.equal(await page.locator('#errorText').textContent(), expected);
    assert.equal(await page.locator('#downloadSection').isVisible(), false);
    assert.equal(await page.locator('#convertBtn').isVisible(), true);
    // app.js logs the underlying parse failure
    assertClean(problems, { allowConsoleErrors: [/^Conversion error:/, /^Error parsing DOCX:/] });
});

test('a .docx whose part declares more than the size limit is refused, and the next file converts', async t => {
    const { page, problems } = await openPage(t);

    // word/document.xml claims to decompress to 200 MiB, past the 100 MiB per-part limit
    const oversized = declareSize(await plainDocx(), 'word/document.xml', 200 * 1024 * 1024);
    await chooseFile(page, 'buyuk.docx', oversized);
    await page.locator('#convertBtn').click();

    await page.locator('#errorMessage').waitFor({ state: 'visible' });
    const expected = await page.evaluate(() => I18N_STRINGS.tr['error.tooLarge']);
    assert.ok(expected, 'the dictionary has an error.tooLarge message');
    assert.equal(await page.locator('#errorText').textContent(), expected);
    assert.equal(await page.locator('#downloadSection').isVisible(), false);

    await convertFile(page, 'dilekce.docx', await plainDocx());
    assert.equal(await page.locator('#errorMessage').isVisible(), false);
    assertClean(problems, { allowConsoleErrors: [/^Conversion error:/, /^Error parsing DOCX:/] });
});

test('the TR | EN toggle and ?lang=en switch the page to English', async t => {
    const { page, problems } = await openPage(t);
    const english = await page.evaluate(() => I18N_STRINGS.en['convert.button']);
    assert.equal(await page.evaluate(() => document.documentElement.lang), 'tr');
    assert.notEqual(await page.locator('#convertBtn').textContent(), english);

    await page.locator('[data-set-lang="en"]').click();
    assert.equal(await page.locator('#convertBtn').textContent(), english);
    assert.equal(await page.evaluate(() => document.documentElement.lang), 'en');
    assert.equal(await page.locator('[data-set-lang="en"]').getAttribute('aria-pressed'), 'true');
    assertClean(problems);

    const fresh = await openPage(t, '?lang=en');
    assert.equal(await fresh.page.evaluate(() => document.documentElement.lang), 'en');
    assert.equal(await fresh.page.locator('#convertBtn').textContent(), english);
    assertClean(fresh.problems);
});

test('"convert another" brings back the upload zone with no file selected', async t => {
    const { page, problems } = await openPage(t);
    await convertFile(page, 'dilekce.docx', await plainDocx());

    await page.locator('#convertAnother').click();
    await page.locator('#uploadZone').waitFor({ state: 'visible' });
    assert.equal(await page.locator('#downloadSection').isVisible(), false);
    assert.equal(await page.locator('#fileInfo').isVisible(), false);
    assert.equal(await page.locator('#convertBtn').isVisible(), true);
    assert.equal(await page.locator('#convertBtn').isDisabled(), true);
    assert.equal(await page.locator('#fileInput').evaluate(input => input.files.length), 0);
    assertClean(problems);
});
