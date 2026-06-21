import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('render nhiều worksheet với cross sheet placeholder, formula và block', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-multi-sheet-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const summary = template.addWorksheet('Summary');
  const details = template.addWorksheet('Details');

  summary.getCell('A1').value = '{{teacher.name}}';
  summary.getCell('B1').value = { formula: 'SUM(Details!B3:B4)', result: 300 } as ExcelJS.CellValue;
  summary.getCell('A2').value = '{{contract.code}}';

  details.getCell('A1').value = '{{teacher.name}}';
  details.getCell('A2').value = '{{#block contracts}}';
  details.getCell('A3').value = '{{code}}';
  details.getCell('B3').value = '{{amount}}';
  details.getCell('C3').value = { formula: 'Summary!$B$1+B3', result: 400 } as ExcelJS.CellValue;
  details.getCell('A4').value = '{{/block}}';
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    teacher: { name: 'Nguyễn Văn A' },
    contract: { code: 'HD-2026' },
    contracts: [
      { code: 'HD01', amount: 100 },
      { code: 'HD02', amount: 250 },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const renderedSummary = output.getWorksheet('Summary');
  const renderedDetails = output.getWorksheet('Details');
  assert.ok(renderedSummary);
  assert.ok(renderedDetails);

  assert.equal(renderedSummary.getCell('A1').value, 'Nguyễn Văn A');
  assert.equal(renderedSummary.getCell('A2').value, 'HD-2026');
  assert.deepEqual(renderedSummary.getCell('B1').value, { formula: 'SUM(Details!B3:B4)', result: 300 });

  assert.equal(renderedDetails.getCell('A1').value, 'Nguyễn Văn A');
  assert.equal(renderedDetails.getCell('A2').value, '');
  assert.deepEqual([
    renderedDetails.getCell('A3').value,
    renderedDetails.getCell('B3').value,
    renderedDetails.getCell('A4').value,
    renderedDetails.getCell('B4').value,
    renderedDetails.getCell('A5').value,
  ], ['HD01', 100, 'HD02', 250, '']);
  assert.deepEqual(renderedDetails.getCell('C3').value, { formula: 'Summary!$B$1+B3', result: 400 });
  assert.deepEqual(renderedDetails.getCell('C4').value, { formula: 'Summary!$B$1+B4' });
});
