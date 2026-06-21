import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import test from 'node:test';
import { MergeRange } from '../src/core/merge/merge-range.js';
import { MergeConflictError, MergeTracker } from '../src/core/merge/merge-tracker.js';
import { ExcelJsMergeManager } from '../src/infrastructure/exceljs/excel-js-merge-manager.js';

test('MergeRange parse, shift và detect intersection', () => {
  const range = MergeRange.fromA1('A1:D1', 'Sheet1');
  const shifted = range.shift(4, 0);

  assert.equal(range.toA1(), 'A1:D1');
  assert.equal(shifted.toA1(), 'A5:D5');
  assert.equal(range.intersects(MergeRange.fromA1('C1:F1', 'Sheet1')), true);
  assert.equal(range.intersects(MergeRange.fromA1('A2:D2', 'Sheet1')), false);
});

test('MergeTracker phát hiện merge conflict', () => {
  const tracker = new MergeTracker([
    MergeRange.fromA1('A1:D1', 'Sheet1'),
  ]);

  assert.throws(
    () => tracker.add(MergeRange.fromA1('C1:F1', 'Sheet1')),
    MergeConflictError,
  );

  tracker.add(MergeRange.fromA1('A2:D2', 'Sheet1'));
  assert.equal(tracker.list().length, 2);
});

test('ExcelJsMergeManager đọc merge ranges từ worksheet', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.mergeCells('A1:D1');
  worksheet.mergeCells('A5:F5');

  const manager = new ExcelJsMergeManager(workbook);
  const ranges = await manager.collect('Sheet1');

  assert.deepEqual(ranges.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'A1:D1',
    'A5:F5',
  ]);
});

test('ExcelJsMergeManager clone row merge và validate không chồng lấn', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.mergeCells('A1:D1');

  const manager = new ExcelJsMergeManager(workbook);
  const clones = await manager.cloneForOperation({
    id: 'clone-row',
    type: 'CloneRow',
    sheetName: 'Sheet1',
    sourceRow: 1,
    targetRow: 2,
    count: 2,
  });

  assert.deepEqual(clones.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'A2:D2',
    'A3:D3',
  ]);

  await manager.validateAndApply('Sheet1', clones);
  assert.equal(worksheet.hasMerges, true);
});

test('ExcelJsMergeManager clone block merge theo hướng down', async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1').mergeCells('B2:D3');

  const manager = new ExcelJsMergeManager(workbook);
  const clones = await manager.cloneForOperation({
    id: 'clone-block',
    type: 'CloneBlock',
    sheetName: 'Sheet1',
    sourceRange: MergeRange.fromA1('A1:E4', 'Sheet1').toRangeAddress(),
    targetTopLeft: { sheetName: 'Sheet1', row: 5, column: 1 },
    count: 2,
    direction: 'down',
  });

  assert.deepEqual(clones.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'B6:D7',
    'B10:D11',
  ]);
});

test('ExcelJsMergeManager validateAndApply không tạo merge chồng lấn', async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1').mergeCells('A1:D1');

  const manager = new ExcelJsMergeManager(workbook);

  await assert.rejects(
    () => manager.validateAndApply('Sheet1', [
      MergeRange.fromA1('C1:F1', 'Sheet1').toRangeAddress(),
    ]),
    MergeConflictError,
  );
});
