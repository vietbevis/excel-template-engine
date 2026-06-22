import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

test('generic engine utilities do not contain showcase-specific report concepts', async () => {
  const source = await readFile('src/application/data/column-tree-compiler.ts', 'utf8');
  const forbiddenTerms = [
    'amountInWords',
    'summary',
    'staticTotals',
    'metricTotals',
    'departments',
    'lecturers',
    'OvertimeReport',
  ];

  assert.deepEqual(forbiddenTerms.filter((term) => source.includes(term)), []);
});
