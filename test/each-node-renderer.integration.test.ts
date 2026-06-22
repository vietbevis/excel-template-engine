import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('EachNode render workbook rows with loop metadata and inline nested loops', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-node-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{#each students}}{{index}}:{{name}}:{{first}}:{{last}}{{/each}}';
  const inline = template.addWorksheet('Inline');
  inline.getCell('C5').value = 'Nested: {{#each classes}}{{index}}[{{#each students}}{{parent.index}}.{{index}}:{{name}};{{/each}}]{{/each}}';
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    students: [
      { name: 'A' },
      { name: 'B' },
    ],
    classes: [
      { students: [{ name: 'C' }, { name: 'D' }] },
      { students: [{ name: 'E' }] },
    ],
  });

  assert.deepEqual(result.warnings, []);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');
  const renderedInline = output.getWorksheet('Inline');
  assert.ok(rendered);
  assert.ok(renderedInline);

  assert.equal(rendered.getCell('A1').value, '0:A:true:false');
  assert.equal(rendered.getCell('A2').value, '1:B:false:true');
  assert.equal(renderedInline.getCell('C5').value, 'Nested: 0[0.0:C;0.1:D;]1[1.0:E;]');
});

test('EachNode clone đầy đủ style từ ô template sang các hàng được render', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-each-node-full-style-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Rows');
  worksheet.getCell('B2').value = '{{#each students}}{{score}}{{/each}}';
  worksheet.getCell('B2').style = {
    numFmt: '0.00',
    font: { name: 'Arial', italic: true, color: { argb: 'FF0F766E' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFBF1' } },
    border: {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'thin' },
    },
    alignment: { horizontal: 'right', vertical: 'middle' },
  };
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    students: [
      { score: 8.5 },
      { score: 9.25 },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Rows');
  assert.ok(rendered);

  assert.equal(rendered.getCell('B2').value, '8.5');
  assert.equal(rendered.getCell('B3').value, '9.25');
  assert.deepEqual(rendered.getCell('B3').style, rendered.getCell('B2').style);
});
