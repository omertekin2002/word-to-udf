/**
 * Test helpers for the page itself: load index.html in jsdom and run the site's
 * scripts in it, the way its <script> tags do.
 */

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const JSZip = require('jszip');
const { JSDOM, VirtualConsole } = require('jsdom');

const ROOT = path.resolve(__dirname, '..', '..');
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

// The app reads files through JSZip, which takes a Buffer in Node; give it the File fields app.js uses
const fileOf = (name, bytes, lastModified) =>
    Object.assign(Buffer.from(bytes), { name, size: bytes.length }, lastModified === undefined ? {} : { lastModified });

const visible = el => !el.classList.contains('hidden');

module.exports = { PAGE_URL, loadPage, fileOf, visible };
