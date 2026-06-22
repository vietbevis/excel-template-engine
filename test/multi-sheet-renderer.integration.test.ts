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
  summary.getCell('B1').value = { formula: 'SUM(Details!B2:B3)', result: 300 } as ExcelJS.CellValue;
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
  assert.deepEqual(renderedSummary.getCell('B1').value, { formula: 'SUM(Details!B2:B3)', result: 300 });

  assert.equal(renderedDetails.getCell('A1').value, 'Nguyễn Văn A');
  assert.deepEqual([
    renderedDetails.getCell('A2').value,
    renderedDetails.getCell('B2').value,
    renderedDetails.getCell('A3').value,
    renderedDetails.getCell('B3').value,
    renderedDetails.getCell('A4').value,
  ], ['HD01', 100, 'HD02', 250, '']);
  assert.deepEqual(renderedDetails.getCell('C2').value, { formula: 'Summary!$B$1+B2', result: 400 });
  assert.deepEqual(renderedDetails.getCell('C3').value, { formula: 'Summary!$B$1+B3' });
});

test('clone worksheet template và render mỗi sheet bằng data scope riêng', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-sheet-clone-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const summary = template.addWorksheet('Summary');
  const departmentTemplate = template.addWorksheet('DepartmentTemplate');

  summary.getCell('A1').value = '{{title}}';
  summary.getCell('A2').value = { formula: "'KHOA CNTT'!B2+'KHOA MẬT MÃ'!B2" } as ExcelJS.CellValue;
  departmentTemplate.getColumn(1).width = 20;
  departmentTemplate.getCell('A1').value = '{{name}}';
  departmentTemplate.getCell('A1').font = { bold: true };
  departmentTemplate.getCell('A2').value = '{{#block lecturers}}';
  departmentTemplate.getCell('A3').value = '{{name}}';
  departmentTemplate.getCell('B3').value = '{{hours}}';
  departmentTemplate.getCell('A3').border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };
  departmentTemplate.getCell('B3').border = departmentTemplate.getCell('A3').border;
  departmentTemplate.getCell('A4').value = '{{/block}}';
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    title: 'Tổng hợp',
    departments: [
      {
        name: 'KHOA CNTT',
        lecturers: [
          { name: 'Nguyễn Văn A', hours: 120 },
          { name: 'Trần Thị B', hours: 80 },
        ],
      },
      {
        name: 'KHOA MẬT MÃ',
        lecturers: [
          { name: 'Lê Văn C', hours: 60 },
        ],
      },
    ],
    _workbook: {
      worksheets: [
        {
          sourceName: 'DepartmentTemplate',
          name: 'KHOA CNTT',
          dataPath: 'departments[0]',
          deleteSource: true,
        },
        {
          sourceName: 'DepartmentTemplate',
          name: 'KHOA MẬT MÃ',
          dataPath: 'departments[1]',
          deleteSource: true,
        },
      ],
    },
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);

  assert.ok(output.getWorksheet('Summary'));
  assert.equal(output.getWorksheet('DepartmentTemplate'), undefined);

  const renderedSummary = output.getWorksheet('Summary');
  const cntt = output.getWorksheet('KHOA CNTT');
  const matMa = output.getWorksheet('KHOA MẬT MÃ');
  assert.ok(renderedSummary);
  assert.ok(cntt);
  assert.ok(matMa);

  assert.equal(renderedSummary.getCell('A1').value, 'Tổng hợp');
  assert.deepEqual(renderedSummary.getCell('A2').value, { formula: "'KHOA CNTT'!B2+'KHOA MẬT MÃ'!B2" });
  assert.equal(cntt.getCell('A1').value, 'KHOA CNTT');
  assert.equal(cntt.getColumn(1).width, 20);
  assert.equal(cntt.getCell('A1').font?.bold, true);
  assert.equal(cntt.getCell('A1').note, undefined);
  assert.deepEqual([cntt.getCell('A2').value, cntt.getCell('B2').value], ['Nguyễn Văn A', 120]);
  assert.deepEqual([cntt.getCell('A3').value, cntt.getCell('B3').value], ['Trần Thị B', 80]);
  assert.equal(cntt.getCell('A2').border?.top?.style, 'thin');
  assert.equal(matMa.getCell('A1').value, 'KHOA MẬT MÃ');
  assert.deepEqual([matMa.getCell('A2').value, matMa.getCell('B2').value], ['Lê Văn C', 60]);
});
