import ExcelJS from 'exceljs';
import { performance } from 'node:perf_hooks';
import { ExcelJsStyleCloneManager } from '../src/infrastructure/exceljs/excel-js-style-clone-manager.js';

const iterations = 5_000;
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');
const manager = new ExcelJsStyleCloneManager(workbook);

worksheet.getCell('A1').style = {
  font: { name: 'Arial', bold: true, color: { argb: 'FFFF0000' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
  border: { top: { style: 'thin', color: { argb: 'FF000000' } } },
  alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
  numFmt: '#,##0.00',
  protection: { locked: true },
};

const sourceRow = worksheet.getRow(1);
sourceRow.height = 32;
sourceRow.font = { name: 'Calibri', bold: true };
sourceRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
sourceRow.border = { bottom: { style: 'medium' } };
sourceRow.alignment = { horizontal: 'right' };
sourceRow.numFmt = '0%';
sourceRow.protection = { locked: false };

const sourceColumn = worksheet.getColumn(1);
sourceColumn.width = 24;
sourceColumn.font = { name: 'Times New Roman', italic: true };
sourceColumn.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000FF' } };
sourceColumn.border = { left: { style: 'dotted' } };
sourceColumn.alignment = { vertical: 'middle' };
sourceColumn.numFmt = '@';
sourceColumn.protection = { locked: true };

const cellStart = performance.now();
for (let index = 0; index < iterations; index += 1) {
  await manager.cloneCellStyle(
    { sheetName: 'Sheet1', row: 1, column: 1 },
    { sheetName: 'Sheet1', row: 1, column: 2 + index },
  );
}
const cellDuration = performance.now() - cellStart;

const rowStart = performance.now();
for (let index = 0; index < iterations; index += 1) {
  await manager.cloneRowStyle('Sheet1', 1, 2 + index);
}
const rowDuration = performance.now() - rowStart;

const columnStart = performance.now();
for (let index = 0; index < iterations; index += 1) {
  await manager.cloneColumnStyle('Sheet1', 1, 2 + index);
}
const columnDuration = performance.now() - columnStart;

console.log(JSON.stringify({
  iterations,
  cell: {
    durationMs: Math.round(cellDuration),
    opsPerSecond: Math.round(iterations / (cellDuration / 1000)),
  },
  row: {
    durationMs: Math.round(rowDuration),
    opsPerSecond: Math.round(iterations / (rowDuration / 1000)),
  },
  column: {
    durationMs: Math.round(columnDuration),
    opsPerSecond: Math.round(iterations / (columnDuration / 1000)),
  },
}, null, 2));
