import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import test from 'node:test';
import { ExcelJsStyleCloneManager } from '../src/infrastructure/exceljs/excel-js-style-clone-manager.js';

test('ExcelJsStyleCloneManager cloneCellStyle đầy đủ và không mutate style gốc', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  const manager = new ExcelJsStyleCloneManager(workbook);
  const source = worksheet.getCell('A1');
  const target = worksheet.getCell('B1');

  source.style = {
    font: { name: 'Arial', bold: true, color: { argb: 'FFFF0000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
    border: { top: { style: 'thin', color: { argb: 'FF000000' } } },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    numFmt: '#,##0.00',
    protection: { locked: true },
  };

  await manager.cloneCellStyle(
    { sheetName: 'Sheet1', row: 1, column: 1 },
    { sheetName: 'Sheet1', row: 1, column: 2 },
  );

  assert.deepEqual(target.style, source.style);
  assert.notEqual(target.style, source.style);
  assert.notEqual(target.style.font, source.style.font);
  assert.notEqual(target.style.fill, source.style.fill);
  assert.notEqual(target.style.border, source.style.border);
  assert.notEqual(target.style.alignment, source.style.alignment);
  assert.notEqual(target.style.protection, source.style.protection);

  target.font = { ...target.font, bold: false };

  assert.equal(source.font.bold, true);
  assert.equal(target.font.bold, false);
});

test('ExcelJsStyleCloneManager cloneRowStyle clone height và style không mutate source', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  const manager = new ExcelJsStyleCloneManager(workbook);
  const source = worksheet.getRow(1);
  const target = worksheet.getRow(2);

  source.height = 32;
  source.font = { name: 'Calibri', bold: true };
  source.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
  source.border = { bottom: { style: 'medium' } };
  source.alignment = { horizontal: 'right' };
  source.numFmt = '0%';
  source.protection = { locked: false };

  await manager.cloneRowStyle('Sheet1', 1, 2);

  assert.equal(target.height, 32);
  assert.deepEqual(target.font, source.font);
  assert.deepEqual(target.fill, source.fill);
  assert.deepEqual(target.border, source.border);
  assert.deepEqual(target.alignment, source.alignment);
  assert.equal(target.numFmt, source.numFmt);
  assert.deepEqual(target.protection, source.protection);
  assert.notEqual(target.font, source.font);

  target.font = { ...target.font, bold: false };

  assert.equal(source.font.bold, true);
  assert.equal(target.font.bold, false);
});

test('ExcelJsStyleCloneManager cloneColumnStyle clone width và style không mutate source', async () => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  const manager = new ExcelJsStyleCloneManager(workbook);
  const source = worksheet.getColumn(1);
  const target = worksheet.getColumn(2);

  source.width = 24;
  source.font = { name: 'Times New Roman', italic: true };
  source.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } };
  source.border = { left: { style: 'dotted' } };
  source.alignment = { vertical: 'middle' };
  source.numFmt = '@';
  source.protection = { locked: true };

  await manager.cloneColumnStyle('Sheet1', 1, 2);

  assert.equal(target.width, 24);
  assert.deepEqual(target.font, source.font);
  assert.deepEqual(target.fill, source.fill);
  assert.deepEqual(target.border, source.border);
  assert.deepEqual(target.alignment, source.alignment);
  assert.equal(target.numFmt, source.numFmt);
  assert.deepEqual(target.protection, source.protection);
  assert.notEqual(target.font, source.font);

  target.font = { ...target.font, italic: false };

  assert.equal(source.font.italic, true);
  assert.equal(target.font.italic, false);
});
