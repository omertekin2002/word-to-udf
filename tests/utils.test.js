const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const context = vm.createContext({});
vm.runInContext(fs.readFileSync(path.resolve(__dirname, '..', 'js', 'utils.js'), 'utf8'), context,
    { filename: 'js/utils.js' });
const { toUdfFilename } = context;

test('toUdfFilename replaces only a trailing .docx, in any case', () => {
    const cases = {
        'a.docx': 'a.udf',
        'A.DOCX': 'A.udf',
        'a.Docx': 'a.udf',
        'taslak.docx.son.docx': 'taslak.docx.son.udf',
        'dilekçe 2026.docx': 'dilekçe 2026.udf',
        '.docx': 'document.udf',
        '.DOCX': 'document.udf'
    };
    for (const [input, expected] of Object.entries(cases)) {
        assert.equal(toUdfFilename(input), expected, input);
    }
});

/** Load utils.js with the given globals, e.g. a fake requestAnimationFrame */
function utilsWith(globals) {
    const sandbox = vm.createContext({ setTimeout, ...globals });
    vm.runInContext(fs.readFileSync(path.resolve(__dirname, '..', 'js', 'utils.js'), 'utf8'), sandbox,
        { filename: 'js/utils.js' });
    return sandbox;
}

test('nextFrame continues after the next animation frame, not before it', async () => {
    const frames = [];
    const { nextFrame } = utilsWith({ requestAnimationFrame: callback => frames.push(callback) });
    let resolved = false;
    const waiting = nextFrame().then(() => { resolved = true; });

    await new Promise(resolve => setTimeout(resolve, 10));
    assert.equal(resolved, false, 'still waiting for the frame');
    assert.equal(frames.length, 1);
    frames[0](0);
    await waiting;
    assert.equal(resolved, true);
});

test('nextFrame falls back to a timer without requestAnimationFrame, or when frames are paused', async () => {
    const { nextFrame: withoutFrames } = utilsWith({});
    await withoutFrames();

    // A hidden tab runs no animation frames; the conversion must not stall there
    const paused = utilsWith({ requestAnimationFrame: () => {} });
    const started = Date.now();
    await paused.nextFrame();
    assert.ok(Date.now() - started >= vm.runInContext('NEXT_FRAME_FALLBACK_MS', paused) - 5);
});
