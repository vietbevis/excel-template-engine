import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { MergeRange } from '../src/core/merge/merge-range.js';
import { ExcelTemplateEngine } from '../src/index.js';
import { ExcelJsMergeManager } from '../src/infrastructure/exceljs/excel-js-merge-manager.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('EachColumnNode render dynamic columns và clone width/style/merge', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getColumn(1).width = 24;
  worksheet.getCell('A1').value = '{{#each-col subjects}}{{name}}{{/each-col}}';
  worksheet.getCell('A1').style = {
    font: { name: 'Arial', bold: true, color: { argb: 'FFFF0000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
    border: { top: { style: 'thin', color: { argb: 'FF000000' } } },
    alignment: { horizontal: 'center', vertical: 'middle' },
    numFmt: '@',
    protection: { locked: true },
  };
  worksheet.mergeCells('A1:A2');
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    subjects: [
      { name: 'Toán' },
      { name: 'Lý' },
      { name: 'Hóa' },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');

  assert.equal(rendered.getCell('A1').value, 'Toán');
  assert.equal(rendered.getCell('B1').value, 'Lý');
  assert.equal(rendered.getCell('C1').value, 'Hóa');

  assert.equal(rendered.getColumn(1).width, 24);
  assert.equal(rendered.getColumn(2).width, 24);
  assert.equal(rendered.getColumn(3).width, 24);

  assert.deepEqual(rendered.getCell('B1').font, rendered.getCell('A1').font);
  assert.deepEqual(rendered.getCell('C1').fill, rendered.getCell('A1').fill);
  assert.deepEqual(rendered.getCell('B1').border, rendered.getCell('A1').border);
  assert.deepEqual(rendered.getCell('C1').alignment, rendered.getCell('A1').alignment);
  assert.equal(rendered.getCell('B1').numFmt, rendered.getCell('A1').numFmt);
  assert.deepEqual(rendered.getCell('C1').protection, rendered.getCell('A1').protection);

  const mergeManager = new ExcelJsMergeManager(output);
  const merges = await mergeManager.collect('Report');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'A1:A2',
    'B1:B2',
    'C1:C2',
  ]);
});

test('EachColumnNode clone border cho styled empty cells trong source column', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-empty-border-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('ColumnBorders');
  worksheet.getCell('A1').value = '{{#each-col columns}}{{name}}{{/each-col}}';
  const border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  } satisfies Partial<ExcelJS.Borders>;
  for (const address of ['A1', 'A2', 'A3']) {
    worksheet.getCell(address).border = border;
  }
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    columns: [
      { name: 'A' },
      { name: 'B' },
      { name: 'C' },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('ColumnBorders');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 'A');
  assert.equal(rendered.getCell('B1').value, 'B');
  assert.equal(rendered.getCell('C1').value, 'C');
  for (const address of ['B2', 'B3', 'C2', 'C3']) {
    assert.deepEqual(rendered.getCell(address).border, border);
  }
});

test('EachColumnNode clone đầy đủ style từ ô template sang các ô được render', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-full-style-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('FullStyle');
  worksheet.getCell('A1').value = '{{#each-col columns reserve=3}}{{amount}}{{/each-col}}';
  worksheet.getCell('A1').style = {
    numFmt: '#,##0.00',
    font: { name: 'Arial', bold: true, color: { argb: 'FF1F2937' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } },
    border: {
      top: { style: 'thin', color: { argb: 'FF111827' } },
      left: { style: 'thin', color: { argb: 'FF111827' } },
      bottom: { style: 'medium', color: { argb: 'FF111827' } },
      right: { style: 'thin', color: { argb: 'FF111827' } },
    },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    protection: { locked: false },
  };
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    columns: [
      { amount: 10 },
      { amount: 20 },
      { amount: 30 },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('FullStyle');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('A1').value,
    rendered.getCell('B1').value,
    rendered.getCell('C1').value,
  ], [10, 20, 30]);
  assert.deepEqual(rendered.getCell('B1').style, rendered.getCell('A1').style);
  assert.deepEqual(rendered.getCell('C1').style, rendered.getCell('A1').style);
});

test('EachColumnNode hỗ trợ span, rowspan và block body dynamic columns', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-span-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('DynamicHeader');
  worksheet.mergeCells('A1:I1');
  worksheet.getCell('A1').value = 'Dynamic Grouped Header Test';
  worksheet.mergeCells('A2:A3');
  worksheet.getCell('A2').value = 'Category';
  worksheet.getCell('B2').value = '{{#each-col groups span=size}}{{name}}{{/each-col}}';
  worksheet.getCell('B3').value = '{{#each-col groupCols}}{{col}}{{/each-col}}';
  worksheet.getCell('A4').value = '{{#block rows}}';
  worksheet.getCell('A5').value = '{{label}}';
  worksheet.getCell('B5').value = '{{#each-col rowCols}}{{value}}{{/each-col}}';
  worksheet.getCell('A6').value = '{{/block}}';

  for (const address of ['A2', 'B2', 'B3', 'A5', 'B5']) {
    worksheet.getCell(address).style = {
      font: { name: 'Arial', bold: true },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } },
      border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } },
      alignment: { horizontal: 'center', vertical: 'middle' },
    };
  }
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    groups: [
      { name: 'Nhóm A', size: 3 },
      { name: 'Nhóm B', size: 2 },
      { name: 'Nhóm C', size: 3 },
    ],
    groupCols: [
      { col: 'A1' },
      { col: 'A2' },
      { col: 'A3' },
      { col: 'B1' },
      { col: 'B2' },
      { col: 'C1' },
      { col: 'C2' },
      { col: 'C3' },
    ],
    rows: [
      { label: 'Row 1', rowCols: [1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ value })) },
      { label: 'Row 2', rowCols: [9, 10, 11, 12, 13, 14, 15, 16].map((value) => ({ value })) },
    ],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('DynamicHeader');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('B2').value,
    rendered.getCell('E2').value,
    rendered.getCell('G2').value,
  ], ['Nhóm A', 'Nhóm B', 'Nhóm C']);

  assert.deepEqual([
    rendered.getCell('B3').value,
    rendered.getCell('C3').value,
    rendered.getCell('D3').value,
    rendered.getCell('E3').value,
    rendered.getCell('F3').value,
    rendered.getCell('G3').value,
    rendered.getCell('H3').value,
    rendered.getCell('I3').value,
  ], ['A1', 'A2', 'A3', 'B1', 'B2', 'C1', 'C2', 'C3']);

  assert.deepEqual([
    rendered.getCell('A4').value,
    rendered.getCell('B4').value,
    rendered.getCell('I4').value,
    rendered.getCell('A5').value,
    rendered.getCell('B5').value,
    rendered.getCell('I5').value,
  ], ['Row 1', 1, 8, 'Row 2', 9, 16]);

  assert.deepEqual(rendered.getCell('C2').fill, rendered.getCell('B2').fill);
  assert.deepEqual(rendered.getCell('D2').border, rendered.getCell('B2').border);

  const mergeManager = new ExcelJsMergeManager(output);
  const merges = await mergeManager.collect('DynamicHeader');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()).sort(), [
    'A1:I1',
    'A2:A3',
    'B2:D2',
    'E2:F2',
    'G2:I2',
  ]);
});

test('EachColumnNode reserve=N dùng vùng cột có sẵn thay vì clone thêm', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-reserve-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('ReservedColumns');
  worksheet.getCell('A1').value = 'Static';
  worksheet.getCell('B1').value = '{{#each-col groups span=size reserve=4}}{{name}}{{/each-col}}';
  worksheet.getCell('B2').value = '{{#each-col columns reserve=4}}{{name}}{{/each-col}}';
  worksheet.getCell('F1').value = 'After dynamic area';
  worksheet.getCell('B1').style = {
    font: { name: 'Arial', bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  };
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    groups: [
      { name: 'Group A', size: 2 },
      { name: 'Group B', size: 2 },
    ],
    columns: [
      { name: 'A1' },
      { name: 'A2' },
      { name: 'B1' },
      { name: 'B2' },
    ],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('ReservedColumns');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('B1').value,
    rendered.getCell('D1').value,
    rendered.getCell('B2').value,
    rendered.getCell('C2').value,
    rendered.getCell('D2').value,
    rendered.getCell('E2').value,
    rendered.getCell('F1').value,
  ], ['Group A', 'Group B', 'A1', 'A2', 'B1', 'B2', 'After dynamic area']);

  const merges = await new ExcelJsMergeManager(output).collect('ReservedColumns');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()).sort(), [
    'B1:C1',
    'D1:E1',
  ]);
});

test('EachColumnNode liền kề nhau trong cùng row không ghi đè nhau', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-adjacent-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Adjacent');
  worksheet.getCell('A1').value = '{{#each-col first}}{{name}}{{/each-col}}';
  worksheet.getCell('B1').value = '{{#each-col second}}{{name}}{{/each-col}}';
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    first: [{ name: 'A1' }, { name: 'A2' }, { name: 'A3' }],
    second: [{ name: 'B1' }, { name: 'B2' }],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Adjacent');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('A1').value,
    rendered.getCell('B1').value,
    rendered.getCell('C1').value,
    rendered.getCell('D1').value,
    rendered.getCell('E1').value,
  ], ['A1', 'A2', 'A3', 'B1', 'B2']);
});

test('EachColumnNode giữ formula object từ data', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-formula-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('FormulaValues');
  worksheet.getCell('A1').value = '{{#each-col metrics}}{{value}}{{/each-col}}';
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    metrics: [
      { value: 10 },
      { value: { formula: 'A{row}*2' } },
    ],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('FormulaValues');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 10);
  assert.deepEqual(rendered.getCell('B1').value, { formula: 'A1*2' });
});

test('EachColumnNode render=false consume span nhưng không ghi header phụ', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-col-render-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('RenderFlag');
  worksheet.getCell('A1').value = '{{#each-col bands span=span rowspan=rowSpan reserve=6}}{{name}}{{/each-col}}';
  worksheet.getCell('A2').value = '{{#each-col groups span=span render=renderHeader reserve=6}}{{name}}{{/each-col}}';
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    bands: [
      { name: 'A', span: 2, rowSpan: 1 },
      { name: 'B', span: 3, rowSpan: 2 },
      { name: 'C', span: 1, rowSpan: 1 },
    ],
    groups: [
      { name: 'A1', span: 2, renderHeader: true },
      { name: 'B1', span: 3, renderHeader: false },
      { name: 'C1', span: 1, renderHeader: true },
    ],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('RenderFlag');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A2').value, 'A1');
  assert.equal(rendered.getCell('F2').value, 'C1');

  const merges = await new ExcelJsMergeManager(output).collect('RenderFlag');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()).sort(), [
    'A1:B1',
    'A2:B2',
    'C1:E2',
  ]);
});
