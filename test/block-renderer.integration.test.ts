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

  assert.equal(rendered.getCell('A1').value, '');
  assert.equal(rendered.getCell('A2').value, 'HD01');
  assert.equal(rendered.getCell('C2').value, 100);
  assert.deepEqual(rendered.getCell('D2').value, { formula: 'C2*2', result: 200 });
  assert.equal(rendered.getCell('A3').value, 'HD02');
  assert.equal(rendered.getCell('C3').value, 250);
  assert.deepEqual(rendered.getCell('D3').value, { formula: 'C3*2' });
  assert.equal(rendered.getCell('A4').value, '');

  assert.equal(rendered.getRow(3).height, 24);
  assert.deepEqual(rendered.getCell('A3').font, rendered.getCell('A2').font);
  assert.deepEqual(rendered.getCell('A3').fill, rendered.getCell('A2').fill);
  assert.equal(rendered.getCell('C3').numFmt, '#,##0');
  assert.deepEqual(rendered.getCell('D3').font, rendered.getCell('D2').font);

  const mergeManager = new ExcelJsMergeManager(output);
  const merges = await mergeManager.collect('Report');
  assert.deepEqual(merges.map((range) => MergeRange.fromRangeAddress(range).toA1()), [
    'A2:B2',
    'A3:B3',
  ]);
  assert.equal(rendered.getImages().length, 2);
});
