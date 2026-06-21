import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('ExcelTemplateEngine.registerHelper renders helper calls in workbook cells', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-helper-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{currency(price)}}';
  worksheet.getCell('B1').value = '{{sum(scores)}}';
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  engine
    .registerHelper('currency', ([value]) => `${Number(value).toLocaleString('en-US')} VND`)
    .registerHelper('sum', ([values]) => Array.isArray(values)
      ? values.reduce((total, value) => total + Number(value), 0)
      : 0);

  const result = await engine.render(templatePath, {
    price: 1250000,
    scores: [8, 9, 10],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, '1,250,000 VND');
  assert.equal(rendered.getCell('B1').value, 27);
});
