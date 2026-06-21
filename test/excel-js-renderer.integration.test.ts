import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

test('ExcelJsWorkbookRenderer load template, render AST và export xlsx output', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{teacher.name}}';
  worksheet.getCell('B1').value = '{{contract.code}}';
  worksheet.getCell('C1').value = '{{user.profile.email}}';
  worksheet.getCell('A2').value = '{{teacher.nickname ?? "Unknown"}}';
  worksheet.getCell('B2').value = 'Static text';
  worksheet.getCell('C2').value = 123;
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    teacher: {
      name: 'Nguyễn Văn A',
    },
    contract: {
      code: 'HD001',
    },
    user: {
      profile: {
        email: 'user@example.com',
      },
    },
  });

  assert.ok(result.output.byteLength > 0);

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');

  assert.equal(rendered.getCell('A1').value, 'Nguyễn Văn A');
  assert.equal(rendered.getCell('B1').value, 'HD001');
  assert.equal(rendered.getCell('C1').value, 'user@example.com');
  assert.equal(rendered.getCell('A2').value, 'Unknown');
  assert.equal(rendered.getCell('B2').value, 'Static text');
  assert.equal(rendered.getCell('C2').value, 123);
});
