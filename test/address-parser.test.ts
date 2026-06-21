import test from 'node:test';
import assert from 'node:assert/strict';
import { AddressParser } from '../src/shared/address/address-parser.js';

test('AddressParser đổi tên cột Excel sang số và ngược lại', () => {
  const parser = new AddressParser();

  assert.equal(parser.columnNameToNumber('A'), 1);
  assert.equal(parser.columnNameToNumber('Z'), 26);
  assert.equal(parser.columnNameToNumber('AA'), 27);
  assert.equal(parser.columnNameToNumber('XFD'), 16384);
  assert.equal(parser.columnNumberToName(1), 'A');
  assert.equal(parser.columnNumberToName(27), 'AA');
  assert.equal(parser.columnNumberToName(16384), 'XFD');
});

test('AddressParser parse cell A1 có sheet và absolute reference', () => {
  const parser = new AddressParser();

  assert.deepEqual(parser.parseCell("'Bảng điểm'!$C$12"), {
    sheetName: 'Bảng điểm',
    columnName: 'C',
    column: 3,
    row: 12,
    absoluteColumn: true,
    absoluteRow: true,
  });
});

test('AddressParser parse range và kế thừa sheet name cho điểm cuối', () => {
  const parser = new AddressParser();

  assert.deepEqual(parser.parseRange('Sheet1!A1:D10'), {
    sheetName: 'Sheet1',
    start: {
      sheetName: 'Sheet1',
      row: 1,
      column: 1,
    },
    end: {
      sheetName: 'Sheet1',
      row: 10,
      column: 4,
    },
  });
});
