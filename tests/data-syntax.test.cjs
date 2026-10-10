const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

test('data.js loads all datasets and preserves the multiline market summary', () => {
  const source = fs.readFileSync(path.join(__dirname, '../data.js'), 'utf8');
  const context = vm.createContext({});
  assert.doesNotThrow(() => vm.runInContext(source, context));
  assert.ok(vm.runInContext('DATA_EXPORT.length > 0 && DATA_KR_COUNTRY.length > 0', context));
  const summary = vm.runInContext('DATA_NEWS.items[0].summary', context);
  assert.ok(summary.includes('\n\n**수요:**'));
  assert.ok(summary.includes('4,948 위안/톤'));
});
