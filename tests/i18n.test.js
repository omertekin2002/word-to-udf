const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const JSZip = require('jszip');
const { JSDOM, VirtualConsole } = require('jsdom');
const { buildDocx } = require('./helpers');

const ROOT = path.resolve(__dirname, '..');
const HTML = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const PAGE_URL = 'https://word-to-udf.test/';
const APP_SCRIPTS = ['js/utils.js', 'js/docx-parser.js', 'js/udf-generator.js', 'js/i18n.js', 'js/app.js'];

/**
 * Load index.html and run the page's scripts in it, the way the <script> tags do
 * @param {Object} [options]
 * @param {string} [options.url] - Page URL, e.g. with a ?lang= query
 * @param {Object<string,string>} [options.storage] - localStorage contents before the scripts run
 * @param {boolean} [options.app] - Also load the converter and app.js, not just i18n.js
 */
async function loadPage({ url = PAGE_URL, storage = {}, app = false } = {}) {
    const errors = [];
    const virtualConsole = new VirtualConsole();
    virtualConsole.on('error', (...args) => errors.push(args));
    const dom = new JSDOM(HTML, { url, runScripts: 'outside-only', virtualConsole });
    const { window } = dom;
    for (const [key, value] of Object.entries(storage)) window.localStorage.setItem(key, value);

    // Run each file as a classic script, so top-level declarations share one scope as on the page
    const context = dom.getInternalVMContext();
    const run = code => vm.runInContext(code, context);
    window.JSZip = JSZip;
    for (const file of app ? APP_SCRIPTS : ['js/i18n.js']) {
        vm.runInContext(fs.readFileSync(path.join(ROOT, file), 'utf8'), context, { filename: file });
    }
    if (window.document.readyState === 'loading') {
        await new Promise(resolve => window.document.addEventListener('DOMContentLoaded', resolve, { once: true }));
    }
    if (app && !window.app) run('window.app = new WordToUdfApp()');

    return { window, document: window.document, strings: run('I18N_STRINGS'), errors, run };
}

const normalize = text => text.replace(/\s+/g, ' ').trim();

/**
 * Assert every translated piece of the page shows `lang`'s string
 */
function assertPageIn(document, dictionary, lang) {
    assert.equal(document.documentElement.lang, lang);
    assert.equal(document.title, dictionary['meta.title']);
    for (const el of document.querySelectorAll('[data-i18n]')) {
        assert.equal(el.textContent, dictionary[el.dataset.i18n], `[data-i18n="${el.dataset.i18n}"]`);
    }
    for (const el of document.querySelectorAll('[data-i18n-html]')) {
        assert.equal(el.innerHTML, dictionary[el.dataset.i18nHtml], `[data-i18n-html="${el.dataset.i18nHtml}"]`);
    }
    for (const el of document.querySelectorAll('[data-i18n-attr]')) {
        for (const [attr, key] of attrPairs(el)) {
            assert.equal(el.getAttribute(attr), dictionary[key], `${attr} of [data-i18n-attr] ${key}`);
        }
    }
}

function attrPairs(el) {
    return el.dataset.i18nAttr.split(';').map(pair => pair.split(':').map(part => part.trim()));
}

function pressed(document) {
    return Object.fromEntries([...document.querySelectorAll('[data-set-lang]')]
        .map(button => [button.dataset.setLang, button.getAttribute('aria-pressed')]));
}

const STORAGE_KEY = 'word-to-udf.lang';

test('the tr and en dictionaries have the same keys', async () => {
    const { strings } = await loadPage();
    assert.deepEqual(Object.keys(strings).sort(), ['en', 'tr']);
    assert.deepEqual(Object.keys(strings.en).sort(), Object.keys(strings.tr).sort());
    for (const lang of ['tr', 'en']) {
        for (const [key, value] of Object.entries(strings[lang])) {
            assert.ok(typeof value === 'string' && value.trim() !== '', `${lang}.${key} is empty`);
        }
    }
});

test('every key used by index.html and app.js exists in both dictionaries', async () => {
    const { strings } = await loadPage();
    const { document } = new JSDOM(HTML).window;
    const keys = new Set(['meta.title']);
    for (const el of document.querySelectorAll('[data-i18n]')) keys.add(el.dataset.i18n);
    for (const el of document.querySelectorAll('[data-i18n-html]')) keys.add(el.dataset.i18nHtml);
    for (const el of document.querySelectorAll('[data-i18n-attr]')) {
        for (const [attr, key] of attrPairs(el)) {
            assert.ok(attr && key, `malformed data-i18n-attr "${el.dataset.i18nAttr}"`);
            keys.add(key);
        }
    }
    const appSource = fs.readFileSync(path.join(ROOT, 'js/app.js'), 'utf8');
    const appKeys = [...appSource.matchAll(/'((?:error|progress)\.\w+)'/g)].map(match => match[1]);
    assert.ok(appKeys.length >= 6, 'app.js message keys found');
    for (const key of appKeys) keys.add(key);

    for (const key of keys) {
        assert.ok(key in strings.tr, `missing tr key ${key}`);
        assert.ok(key in strings.en, `missing en key ${key}`);
    }
});

test('the static markup is the Turkish text, so the page reads correctly before scripts run', async () => {
    const { strings } = await loadPage();
    const { document } = new JSDOM(HTML).window;
    assert.equal(document.documentElement.lang, 'tr');
    assert.equal(document.title, strings.tr['meta.title']);
    for (const el of document.querySelectorAll('[data-i18n]')) {
        assert.equal(normalize(el.textContent), strings.tr[el.dataset.i18n], `[data-i18n="${el.dataset.i18n}"]`);
    }
    for (const el of document.querySelectorAll('[data-i18n-html]')) {
        assert.equal(normalize(el.innerHTML), strings.tr[el.dataset.i18nHtml]);
    }
    for (const el of document.querySelectorAll('[data-i18n-attr]')) {
        for (const [attr, key] of attrPairs(el)) assert.equal(el.getAttribute(attr), strings.tr[key]);
    }
});

test('a first visit is in Turkish, even though the browser is English', async () => {
    const { window, document, strings } = await loadPage();
    assert.match(window.navigator.language, /^en/);
    assertPageIn(document, strings.tr, 'tr');
    assert.deepEqual(pressed(document), { tr: 'true', en: 'false' });
    assert.equal(window.localStorage.getItem(STORAGE_KEY), null, 'nothing saved until the user chooses');
});

test('the EN button switches the page to English without a reload, and TR switches back', async () => {
    const { window, document, strings } = await loadPage();
    const before = window.document;

    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(window.document, before);
    assertPageIn(document, strings.en, 'en');
    assert.deepEqual(pressed(document), { tr: 'false', en: 'true' });
    assert.equal(window.localStorage.getItem(STORAGE_KEY), 'en');

    document.querySelector('[data-set-lang="tr"]').click();
    assertPageIn(document, strings.tr, 'tr');
    assert.deepEqual(pressed(document), { tr: 'true', en: 'false' });
    assert.equal(window.localStorage.getItem(STORAGE_KEY), 'tr');
});

test('the toggle is a pair of real buttons in a labelled group', async () => {
    const { document } = await loadPage();
    const group = document.querySelector('.lang-toggle');
    assert.equal(group.getAttribute('role'), 'group');
    assert.equal(group.getAttribute('aria-label'), 'Dil');
    const buttons = [...group.querySelectorAll('[data-set-lang]')];
    assert.deepEqual(buttons.map(b => b.textContent), ['TR', 'EN']);
    for (const button of buttons) {
        assert.equal(button.tagName, 'BUTTON');
        assert.equal(button.type, 'button');
        assert.ok(button.getAttribute('aria-label'));
    }
});

test('a saved choice persists across reloads', async () => {
    const first = await loadPage();
    first.document.querySelector('[data-set-lang="en"]').click();
    const saved = first.window.localStorage.getItem(STORAGE_KEY);

    const reloaded = await loadPage({ storage: { [STORAGE_KEY]: saved } });
    assertPageIn(reloaded.document, reloaded.strings.en, 'en');
    assert.deepEqual(pressed(reloaded.document), { tr: 'false', en: 'true' });
});

test('?lang= overrides the saved choice', async () => {
    const en = await loadPage({ url: `${PAGE_URL}?lang=en`, storage: { [STORAGE_KEY]: 'tr' } });
    assertPageIn(en.document, en.strings.en, 'en');

    const tr = await loadPage({ url: `${PAGE_URL}?lang=tr`, storage: { [STORAGE_KEY]: 'en' } });
    assertPageIn(tr.document, tr.strings.tr, 'tr');
});

test('an unknown ?lang= or saved value falls back to Turkish', async () => {
    const page = await loadPage({ url: `${PAGE_URL}?lang=de`, storage: { [STORAGE_KEY]: 'fr' } });
    assertPageIn(page.document, page.strings.tr, 'tr');
});

test('switching language with ?lang= in the URL keeps the URL in step', async () => {
    const { window, document } = await loadPage({ url: `${PAGE_URL}?lang=tr` });
    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(new URL(window.location.href).searchParams.get('lang'), 'en');
});

test('t() falls back to the key and fills {placeholders}', async () => {
    const { run } = await loadPage();
    assert.equal(run("t('no.such.key')"), 'no.such.key');
    run("I18N_STRINGS.tr['test.greeting'] = 'Merhaba {name}'");
    assert.equal(run("t('test.greeting', { name: 'Dünya' })"), 'Merhaba Dünya');
});

// The app reads files through JSZip, which takes a Buffer in Node; give it the File fields app.js uses
const fileOf = (name, bytes) => Object.assign(Buffer.from(bytes), { name, size: bytes.length });

const visible = el => !el.classList.contains('hidden');

test('a non-.docx file gets a translated error, which follows a language change', async () => {
    const { window, document, strings } = await loadPage({ app: true });
    window.app.processFile(fileOf('notes.txt', 'plain text'));

    const errorText = document.getElementById('errorText');
    assert.ok(visible(document.getElementById('errorMessage')));
    assert.equal(errorText.textContent, strings.tr['error.invalidType']);

    document.querySelector('[data-set-lang="en"]').click();
    assert.equal(errorText.textContent, strings.en['error.invalidType']);
});

test('a corrupt .docx gets a translated error and the raw error still goes to the console', async () => {
    const { window, document, strings, errors } = await loadPage({ url: `${PAGE_URL}?lang=en`, app: true });
    window.app.processFile(fileOf('broken.docx', 'this is not a zip archive'));
    await window.app.handleConvert();

    assert.ok(visible(document.getElementById('errorMessage')));
    assert.equal(document.getElementById('errorText').textContent, strings.en['error.parse']);
    assert.ok(errors.some(args => args.some(arg => arg instanceof Error || /parse|zip/i.test(String(arg)))),
        'the underlying error is logged');

    document.querySelector('[data-set-lang="tr"]').click();
    assert.equal(document.getElementById('errorText').textContent, strings.tr['error.parse']);
});

// A .docx whose word/document.xml stops mid-paragraph
async function truncatedDocx() {
    const zip = await JSZip.loadAsync(await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>'));
    const documentXml = await zip.file('word/document.xml').async('string');
    zip.file('word/document.xml', documentXml.slice(0, documentXml.indexOf('</w:t>')));
    return zip.generateAsync({ type: 'nodebuffer' });
}

// The 8-byte signature and a directory entry naming the EncryptionInfo stream, at offset 512
function compoundFile({ encrypted }) {
    const bytes = Buffer.alloc(1024);
    Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]).copy(bytes);
    const name = encrypted ? 'EncryptionInfo' : 'WordDocument';
    bytes.write(name, 512, 'utf16le');
    bytes.writeUInt16LE((name.length + 1) * 2, 512 + 64);
    bytes[512 + 66] = 2;
    return bytes;
}

const NAMED_FAILURES = [
    ['a damaged .docx', 'error.corrupt', truncatedDocx],
    ['a password-protected .docx', 'error.encrypted', async () => compoundFile({ encrypted: true })],
    ['a .doc renamed to .docx', 'error.legacyDoc', async () => compoundFile({ encrypted: false })]
];

for (const [what, key, build] of NAMED_FAILURES) {
    test(`${what} gets ${key} in the page language and no download`, async () => {
        const { window, document, strings } = await loadPage({ app: true });
        window.app.processFile(fileOf('dilekce.docx', await build()));
        await window.app.handleConvert();

        const errorText = document.getElementById('errorText');
        assert.ok(visible(document.getElementById('errorMessage')));
        assert.ok(!visible(document.getElementById('downloadSection')));
        assert.equal(window.app.convertedBlob, null);
        assert.equal(errorText.textContent, strings.tr[key]);

        document.querySelector('[data-set-lang="en"]').click();
        assert.equal(errorText.textContent, strings.en[key]);
    });
}

test('an upper-case .DOCX name downloads as .udf', async () => {
    const { window } = await loadPage({ app: true });
    window.app.processFile(fileOf('DILEKCE.DOCX', 'not converted here'));

    assert.equal(window.app.outputFilename, 'DILEKCE.udf');
});

for (const lang of ['tr', 'en']) {
    test(`a .docx converts end to end with the page in ${lang}`, async () => {
        const { window, document, strings } = await loadPage({ url: `${PAGE_URL}?lang=${lang}`, app: true });
        const docx = await buildDocx('<w:p><w:r><w:t>Dilekçe metni</w:t></w:r></w:p>');
        window.app.processFile(fileOf('dilekce.docx', docx));
        await window.app.handleConvert();

        assert.ok(!visible(document.getElementById('errorMessage')), document.getElementById('errorText').textContent);
        assert.ok(visible(document.getElementById('downloadSection')));
        assert.equal(window.app.outputFilename, 'dilekce.udf');
        assert.equal(document.getElementById('progressText').textContent, strings[lang]['progress.complete']);
        assertPageIn(document, strings[lang], lang);

        const udf = await JSZip.loadAsync(await window.app.convertedBlob.arrayBuffer());
        assert.match(await udf.file('content.xml').async('string'), /Dilekçe metni/);
    });
}
