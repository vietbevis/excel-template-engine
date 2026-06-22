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

test('ExcelJsWorkbookRenderer giữ richText value từ data', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-rich-text-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('RichText');
  worksheet.getCell('A1').value = '{{title}}';
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    title: {
      richText: [
        { text: 'Năm học ' },
        { text: '2025-2026', font: { bold: true, color: { argb: 'FFFF0000' } } },
      ],
    },
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('RichText');
  assert.ok(rendered);

  assert.deepEqual(rendered.getCell('A1').value, {
    richText: [
      { text: 'Năm học ' },
      { text: '2025-2026', font: { bold: true, color: { argb: 'FFFF0000' } } },
    ],
  });
});

test('ExcelJsWorkbookRenderer hỗ trợ word wrap và choice validation từ cell metadata', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-cell-metadata-'));
  const templatePath = join(dir, 'template.xlsx');

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Metadata');
  worksheet.getCell('A1').value = '{{status}}';
  worksheet.getCell('A2').value = '{{note}}';
  worksheet.getCell('B1').value = '{{#each-col approvals reserve=2}}{{value}}{{/each-col}}';
  worksheet.getCell('D3').value = '{{amount}}';
  await template.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    status: {
      value: 'Approved',
      choices: ['Draft', 'Approved', 'Rejected'],
      wrapText: true,
    },
    note: {
      value: 'Line 1\nLine 2',
      wrapText: true,
    },
    amount: {
      value: 12000000,
      numFmt: '#,##0',
    },
    approvals: [
      {
        value: {
          value: 'Yes',
          choice: {
            values: ['Yes', 'No'],
            allowBlank: false,
            errorTitle: 'Invalid choice',
            error: 'Please choose Yes or No.',
          },
        },
      },
      {
        value: {
          value: 'No',
          choice: {
            formula: '$D$1:$D$2',
          },
        },
      },
    ],
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Metadata');
  assert.ok(rendered);

  assert.equal(rendered.getCell('A1').value, 'Approved');
  assert.equal(rendered.getCell('A1').alignment.wrapText, true);
  assert.deepEqual(rendered.getCell('A1').dataValidation, {
    type: 'list',
    allowBlank: true,
    formulae: ['"Draft,Approved,Rejected"'],
    showErrorMessage: true,
  });

  assert.equal(rendered.getCell('A2').value, 'Line 1\nLine 2');
  assert.equal(rendered.getCell('A2').alignment.wrapText, true);
  assert.equal(rendered.getCell('D3').value, 12000000);
  assert.equal(rendered.getCell('D3').numFmt, '#,##0');

  assert.equal(rendered.getCell('B1').value, 'Yes');
  const yesNoValidation = rendered.getCell('B1').dataValidation;
  assert.equal(yesNoValidation.allowBlank ?? false, false);
  assert.deepEqual(yesNoValidation, {
    type: 'list',
    formulae: ['"Yes,No"'],
    showErrorMessage: true,
    errorTitle: 'Invalid choice',
    error: 'Please choose Yes or No.',
  });

  assert.equal(rendered.getCell('C1').value, 'No');
  assert.deepEqual(rendered.getCell('C1').dataValidation, {
    type: 'list',
    allowBlank: true,
    formulae: ['$D$1:$D$2'],
    showErrorMessage: true,
  });
});
