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
