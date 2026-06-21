import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('GridNode sinh cột, sinh hàng và fill dữ liệu vào ô giao nhau', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-grid-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = 'Sinh viên';
  worksheet.getCell('B1').value = '{{#each-col subjects}}{{name}}{{/each-col}}';
  worksheet.getCell('A2').value = '{{#each students}}{{name}}{{/each}}';
  worksheet.getCell('B2').value = '{{#grid students subjects}}{{score}}{{/grid}}';
  worksheet.getColumn(2).width = 16;
  worksheet.getRow(2).height = 28;
  worksheet.getCell('B2').style = {
    font: { name: 'Arial', bold: true },
    alignment: { horizontal: 'center' },
    numFmt: '0',
  };
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    subjects: [
      { name: 'Toán' },
      { name: 'Lý' },
      { name: 'Hóa' },
    ],
    students: [
      { name: 'A', scores: [8, 9, 10] },
      { name: 'B', scores: [7, 6, 8] },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');

  assert.deepEqual([
    rendered.getCell('A1').value,
    rendered.getCell('B1').value,
    rendered.getCell('C1').value,
    rendered.getCell('D1').value,
  ], ['Sinh viên', 'Toán', 'Lý', 'Hóa']);

  assert.deepEqual([
    rendered.getCell('A2').value,
    rendered.getCell('B2').value,
    rendered.getCell('C2').value,
    rendered.getCell('D2').value,
  ], ['A', '8', '9', '10']);

  assert.deepEqual([
    rendered.getCell('A3').value,
    rendered.getCell('B3').value,
    rendered.getCell('C3').value,
    rendered.getCell('D3').value,
  ], ['B', '7', '6', '8']);

  assert.equal(rendered.getColumn(3).width, 16);
  assert.equal(rendered.getColumn(4).width, 16);
  assert.equal(rendered.getRow(3).height, 28);
  assert.deepEqual(rendered.getCell('C3').font, rendered.getCell('B2').font);
  assert.deepEqual(rendered.getCell('D3').alignment, rendered.getCell('B2').alignment);
});
