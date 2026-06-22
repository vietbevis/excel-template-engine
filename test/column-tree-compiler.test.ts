import assert from 'node:assert/strict';
import test from 'node:test';
import { ColumnTreeCompiler, FormulaTemplateCompiler } from '../src/index.js';

test('ColumnTreeCompiler derives grouped columns without template-specific fields', () => {
  const result = new ColumnTreeCompiler().compile([
    {
      key: 'teaching',
      name: 'Teaching',
      groups: [
        {
          key: 'semester1',
          name: 'Semester 1',
          columns: [
            { key: 'vn', name: 'VN' },
            { key: 'lao', name: 'Lao' },
          ],
        },
        {
          key: 'semester2',
          name: 'Semester 2',
          columns: [
            { key: 'vn', name: 'VN' },
          ],
        },
        {
          key: 'year',
          name: 'Year',
          derive: {
            type: 'sumSameKey',
            from: ['semester1', 'semester2'],
            total: { key: 'total', name: 'Total' },
          },
        },
      ],
    },
  ], { startColumn: 8 });

  assert.deepEqual(result.bands, [{ name: 'Teaching', colCount: 6, rowSpan: 1 }]);
  assert.deepEqual(result.groups.map((group) => [group.key, group.colCount]), [
    ['semester1', 2],
    ['semester2', 1],
    ['year', 3],
  ]);
  assert.deepEqual(result.columns.map((column) => [column.fullKey, column.columnName]), [
    ['semester1.vn', 'H'],
    ['semester1.lao', 'I'],
    ['semester2.vn', 'J'],
    ['year.vn', 'K'],
    ['year.lao', 'L'],
    ['year.total', 'M'],
  ]);

  const compiler = new FormulaTemplateCompiler(result.columns);
  assert.equal(
    compiler.compile(result.columns.find((column) => column.fullKey === 'year.vn')!.formula!, { row: 3 }),
    'SUM(H$3,J$3)',
  );
  assert.equal(
    compiler.compile(result.columns.find((column) => column.fullKey === 'year.lao')!.formula!, { row: 3 }),
    'SUM(I$3)',
  );
});

test('ColumnTreeCompiler keeps merge rowSpan as input data, not an engine guess', () => {
  const result = new ColumnTreeCompiler().compile([
    {
      key: 'single',
      name: 'Single',
      columns: [{ key: 'value', name: 'Value' }],
    },
    {
      key: 'merged',
      name: 'Merged',
      rowSpan: 3,
      columns: [{ key: 'value', name: 'Value' }],
    },
  ], { startColumn: 1 });

  assert.deepEqual(result.bands, [
    { name: 'Single', colCount: 1, rowSpan: 1 },
    { name: 'Merged', colCount: 1, rowSpan: 3 },
  ]);
});
