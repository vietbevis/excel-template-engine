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
