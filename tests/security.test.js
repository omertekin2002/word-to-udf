const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');

const ROOT = path.resolve(__dirname, '..');
const HTML = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const PAGE_URL = 'https://word-to-udf.test/';
const VENDORED_JSZIP = 'js/vendor/jszip.min.js';
const JSZIP_PACKAGE = path.dirname(require.resolve('jszip/package.json'));

// Parse only: the page's scripts don't run
const { document } = new JSDOM(HTML, { url: PAGE_URL }).window;

/**
 * The page's Content-Security-Policy meta tag, as directive name -> source list
 * @returns {Map<string,string[]>}
 */
function contentSecurityPolicy() {
    const metas = [...document.querySelectorAll('meta[http-equiv]')]
        .filter(meta => meta.getAttribute('http-equiv').toLowerCase() === 'content-security-policy');
    assert.equal(metas.length, 1, 'index.html declares exactly one Content-Security-Policy meta tag');
    const directives = new Map();
    for (const directive of metas[0].getAttribute('content').split(';')) {
        const [name, ...sources] = directive.trim().split(/\s+/);
        if (name) directives.set(name.toLowerCase(), sources);
    }
    return directives;
}

// Files the page ships, other than vendored ones
function siteFiles(dir, extension) {
    return fs.readdirSync(path.join(ROOT, dir), { withFileTypes: true })
        .filter(entry => entry.isFile() && entry.name.endsWith(extension))
        .map(entry => path.join(dir, entry.name));
}

test('the vendored JSZip is byte-identical to the pinned npm package', () => {
    const vendored = fs.readFileSync(path.join(ROOT, VENDORED_JSZIP));
    const packaged = fs.readFileSync(path.join(JSZIP_PACKAGE, 'dist', 'jszip.min.js'));
    assert.ok(vendored.equals(packaged),
        `${VENDORED_JSZIP} differs from node_modules/jszip/dist/jszip.min.js: re-copy it after changing the jszip version`);
});

test('the jszip devDependency is pinned to the exact installed version', () => {
    const { devDependencies } = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
    const { version } = require('jszip/package.json');
    assert.match(devDependencies.jszip, /^\d+\.\d+\.\d+$/, 'jszip is pinned to an exact version, not a range');
    assert.equal(devDependencies.jszip, version);
    assert.ok(fs.readFileSync(path.join(ROOT, VENDORED_JSZIP), 'utf8').includes(`JSZip v${version} `),
        'the vendored file keeps its license header');
});

test('the page loads JSZip from the repo, before the scripts that use it', () => {
    const scripts = [...document.querySelectorAll('script')].map(script => script.getAttribute('src'));
    assert.equal(scripts[0], VENDORED_JSZIP);
    for (const src of scripts) {
        assert.ok(fs.existsSync(path.join(ROOT, src)), `${src} exists in the repo`);
    }
});

test('every subresource of the page comes from its own origin', () => {
    // Links the user follows (<a href>) are navigations, not subresources
    const subresources = [...document.querySelectorAll('[src], [href], [srcset], [data], [poster]')]
        .filter(el => el.localName !== 'a')
        .flatMap(el => ['src', 'href', 'srcset', 'data', 'poster']
            .filter(attr => el.hasAttribute(attr))
            .map(attr => ({ el: el.localName, url: el.getAttribute(attr) })));
    assert.ok(subresources.length > 0);
    for (const { el, url } of subresources) {
        assert.equal(new URL(url, PAGE_URL).origin, new URL(PAGE_URL).origin, `<${el}> loads ${url} from another origin`);
    }
});

test('the stylesheets and scripts reference no other host', () => {
    for (const file of siteFiles('css', '.css')) {
        const css = fs.readFileSync(path.join(ROOT, file), 'utf8');
        assert.doesNotMatch(css, /@import|url\(\s*['"]?(https?:)?\/\//i, `${file} loads nothing from another host`);
    }
    // The only absolute URLs in the scripts are XML namespace names, which are never fetched
    const NAMESPACE_HOSTS = ['schemas.openxmlformats.org', 'schemas.microsoft.com', 'purl.oclc.org'];
    for (const file of siteFiles('js', '.js')) {
        const js = fs.readFileSync(path.join(ROOT, file), 'utf8');
        for (const [url] of js.matchAll(/(https?:)?\/\/[\w.-]+\.[a-z]{2,}[^\s'"`)]*/gi)) {
            const { protocol, hostname } = new URL(url, 'https://example.invalid');
            assert.equal(protocol, 'http:', `${file} references ${url}`);
            assert.ok(NAMESPACE_HOSTS.includes(hostname), `${file} references ${url}`);
        }
        assert.doesNotMatch(js, /\b(fetch|XMLHttpRequest|WebSocket|EventSource|sendBeacon|importScripts)\b/,
            `${file} makes no network requests`);
    }
});

test('the Content-Security-Policy allows only scripts from the page itself', () => {
    const csp = contentSecurityPolicy();
    assert.deepEqual(csp.get('default-src'), ["'none'"]);
    assert.deepEqual(csp.get('script-src'), ["'self'"],
        "script-src allows no 'unsafe-inline', 'unsafe-eval' or remote host");
    assert.deepEqual(csp.get('style-src'), ["'self'"], "style-src allows no 'unsafe-inline' or remote host");
    for (const directive of ['connect-src', 'object-src', 'base-uri', 'form-action']) {
        assert.deepEqual(csp.get(directive), ["'none'"], `${directive} is 'none'`);
    }
    for (const [directive, sources] of csp) {
        for (const source of sources) {
            assert.match(source, /^'(self|none)'$/,
                `${directive} allows only the page's own origin, not ${source}`);
        }
    }
});

test('the Content-Security-Policy comes before anything it governs', () => {
    const meta = document.querySelector('meta[http-equiv]');
    for (const el of document.querySelectorAll('link, script, style')) {
        assert.ok(meta.compareDocumentPosition(el) & document.defaultView.Node.DOCUMENT_POSITION_FOLLOWING,
            `<${el.localName}> comes after the policy`);
    }
});

test('the page has no inline scripts, styles or event handlers the policy would block', () => {
    for (const script of document.querySelectorAll('script')) {
        assert.ok(script.hasAttribute('src'), 'every <script> loads a file');
        assert.equal(script.textContent.trim(), '', `<script src="${script.getAttribute('src')}"> has no inline body`);
    }
    assert.equal(document.querySelectorAll('style').length, 0, 'no <style> elements');
    for (const el of document.querySelectorAll('*')) {
        for (const { name } of el.attributes) {
            assert.notEqual(name.toLowerCase(), 'style', `<${el.localName}> has no inline style attribute`);
            assert.doesNotMatch(name, /^on/i, `<${el.localName}> has no ${name} handler attribute`);
        }
    }
});
