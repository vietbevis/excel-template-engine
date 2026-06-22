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

test('BlockNode clone cells, styles, merge, formula và images', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-block-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{#block contracts}}';
  worksheet.getCell('A2').value = '{{code}}';
  worksheet.getCell('C2').value = '{{amount}}';
  worksheet.getCell('D2').value = { formula: 'C2*2', result: 200 } as ExcelJS.CellValue;
  worksheet.getCell('A3').value = '{{/block}}';
  worksheet.mergeCells('A2:B2');
  worksheet.getRow(2).height = 24;
  worksheet.getCell('A2').style = {
    font: { name: 'Arial', bold: true, color: { argb: 'FF1F2937' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } },
    alignment: { horizontal: 'center' },
  };
  worksheet.getCell('C2').style = {
    numFmt: '#,##0',
    font: { name: 'Arial', italic: true },
  };
  worksheet.getCell('D2').style = {
    font: { name: 'Arial', color: { argb: 'FF047857' } },
  };

  const imageId = template.addImage({
    base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=',
    extension: 'png',
  });
  worksheet.addImage(imageId, {
    tl: { col: 2, row: 1 },
    br: { col: 3, row: 2 },
  });
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    contracts: [
      { code: 'HD01', amount: 100 },
      { code: 'HD02', amount: 250 },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 'HD01');
  assert.equal(rendered.getCell('C1').value, 100);
  assert.deepEqual(rendered.getCell('D1').value, { formula: 'C1*2', result: 200 });
  assert.equal(rendered.getCell('A2').value, 'HD02');
  assert.equal(rendered.getCell('C2').value, 250);
  assert.deepEqual(rendered.getCell('D2').value, { formula: 'C2*2' });
  assert.equal(rendered.getCell('A3').value, '');

  assert.equal(rendered.getRow(2).height, 24);
  assert.deepEqual(rendered.getCell('A2').font, rendered.getCell('A1').font);
  assert.deepEqual(rendered.getCell('A2').fill, rendered.getCell('A1').fill);
  assert.equal(rendered.getCell('C2').numFmt, '#,##0');
  assert.deepEqual(rendered.getCell('D2').font, rendered.getCell('D1').font);

  const mergeManager = new ExcelJsMergeManager(output);
  const merges = await mergeManager.collect('Report');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'A1:B1',
    'A2:B2',
  ]);
  assert.equal(rendered.getImages().length, 2);
});

test('BlockNode clone border cho styled empty cells trong body range', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-block-empty-border-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Borders');
  worksheet.getCell('A1').value = '{{#block rows}}';
  worksheet.getCell('A2').value = '{{name}}';
  worksheet.getCell('A3').value = '{{/block}}';

  const border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  } satisfies Partial<ExcelJS.Borders>;
  for (const address of ['A2', 'B2', 'C2', 'D2']) {
    worksheet.getCell(address).border = border;
  }

  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    rows: [
      { name: 'A' },
      { name: 'B' },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Borders');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 'A');
  assert.equal(rendered.getCell('A2').value, 'B');
  for (const address of ['B1', 'C1', 'D1', 'B2', 'C2', 'D2']) {
    assert.deepEqual(rendered.getCell(address).border, border);
  }
});

test('BlockNode giữ style của row phía sau khi xóa marker rows', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-block-shift-style-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Totals');
  worksheet.getCell('A1').value = '{{#block rows}}';
  worksheet.getCell('A2').value = '{{name}}';
  worksheet.getCell('A3').value = '{{/block}}';
  worksheet.getCell('A4').value = 'Tổng cộng';
  worksheet.getCell('B4').value = '{{total}}';
  for (const address of ['A4', 'B4']) {
    worksheet.getCell(address).font = { name: 'Arial', bold: true };
    worksheet.getCell(address).border = { top: { style: 'thin' } };
  }
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    rows: [
      { name: 'A' },
      { name: 'B' },
    ],
    total: 2,
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Totals');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 'A');
  assert.equal(rendered.getCell('A2').value, 'B');
  assert.equal(rendered.getCell('A3').value, 'Tổng cộng');
  assert.equal(rendered.getCell('B3').value, 2);
  assert.equal(rendered.getCell('A3').font.bold, true);
  assert.equal(rendered.getCell('B3').font.bold, true);
  assert.deepEqual(rendered.getCell('A3').border, { top: { style: 'thin' } });
  assert.deepEqual(rendered.getCell('B3').border, { top: { style: 'thin' } });
});

test('BlockNode hỗ trợ nested block cho nhóm khoa và giảng viên', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-nested-block-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{#block departments}}';
  worksheet.mergeCells('A2:J2');
  worksheet.getCell('A2').value = '{{index}}. {{name}}';
  worksheet.getCell('A2').style = {
    font: { name: 'Arial', bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } },
    alignment: { horizontal: 'center' },
  };
  worksheet.getCell('A3').value = '{{#block lecturers}}';
  worksheet.getCell('A4').value = '{{index}}';
  worksheet.getCell('B4').value = '{{name}}';
  worksheet.getCell('C4').value = '{{#each-col values reserve=valueColumns}}{{value}}{{/each-col}}';
  worksheet.getCell('A5').value = '{{/block}}';
  worksheet.getCell('A6').value = '{{/block}}';
  worksheet.getCell('A7').value = 'Report total: {{reportTotal}}';
  for (const address of ['A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4']) {
    worksheet.getCell(address).style = {
      border: { bottom: { style: 'thin' } },
      alignment: { horizontal: 'center' },
    };
  }
  worksheet.getCell('J4').style = {
    ...worksheet.getCell('J4').style,
    font: { name: 'Arial', bold: true, color: { argb: 'FFB91C1C' } },
  };
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    valueColumns: 8,
    departments: [
      {
        name: 'KHOA CNTT',
        lecturers: [
          { name: 'Nguyen Van A', values: [10, 20, 30, 40, 50, 60, 70, 80].map((value) => ({ value })) },
          { name: 'Tran Thi B', values: [11, 21, 31, 41, 51, 61, 71, 81].map((value) => ({ value })) },
        ],
      },
      {
        name: 'KHOA MAT MA',
        lecturers: [
          { name: 'Le Van C', values: [12, 22, 32, 42, 52, 62, 72, 82].map((value) => ({ value })) },
        ],
      },
    ],
    reportTotal: 3,
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('A1').value,
    rendered.getCell('A2').value,
    rendered.getCell('B2').value,
    rendered.getCell('C2').value,
    rendered.getCell('D2').value,
    rendered.getCell('A3').value,
    rendered.getCell('A4').value,
    rendered.getCell('B5').value,
    rendered.getCell('D5').value,
    rendered.getCell('J5').value,
    rendered.getCell('A6').value,
  ], [
    '0. KHOA CNTT',
    0,
    'Nguyen Van A',
    10,
    20,
    1,
    '1. KHOA MAT MA',
    'Le Van C',
    22,
    82,
    'Report total: 3',
  ]);
  assert.deepEqual(rendered.getCell('J5').font, rendered.getCell('J2').font);
  assert.deepEqual(rendered.getCell('J5').border, rendered.getCell('J2').border);

  rendered.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      assert.equal(/\{\{[^}]+}}/.test(String(cell.value)), false);
    });
  });

  const merges = await new ExcelJsMergeManager(output).collect('Report');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()).sort(), [
    'A1:J1',
    'A4:J4',
  ]);
});
