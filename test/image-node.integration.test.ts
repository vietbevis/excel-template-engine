import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { ExcelTemplateEngine } from '../src/index.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

const pngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';
const jpgBytes = Buffer.from([0xff, 0xd8, 0xff, 0xd9]);

test('ImageNode render png path, jpg path, base64 và buffer sources', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-image-'));
  const templatePath = join(dir, 'template.xlsx');
  await writeFile(join(dir, 'avatar.png'), Buffer.from(pngBase64, 'base64'));
  await writeFile(join(dir, 'avatar.jpg'), jpgBytes);

  const template = new ExcelJS.Workbook();
  const worksheet = template.addWorksheet('Report');
  worksheet.getCell('A1').value = '{{image avatarPath}}';
  worksheet.getCell('B1').value = '{{image avatarBase64}}';
  worksheet.getCell('C1').value = '{{image avatarBuffer}}';
  worksheet.getCell('D1').value = '{{image avatarJpg}}';
  await template.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  const result = await engine.render(templatePath, {
    avatarPath: './avatar.png',
    avatarBase64: pngBase64,
    avatarBuffer: Buffer.from(pngBase64, 'base64'),
    avatarJpg: './avatar.jpg',
  });

  const output = new ExcelJS.Workbook();
  await output.xlsx.load(Buffer.from(result.output) as unknown as ExcelJsLoadInput);
  const rendered = output.getWorksheet('Report');
  assert.ok(rendered);

  assert.deepEqual([
    rendered.getCell('A1').value,
    rendered.getCell('B1').value,
    rendered.getCell('C1').value,
    rendered.getCell('D1').value,
  ], ['', '', '', '']);

  const images = rendered.getImages();
  assert.equal(images.length, 4);
  assert.deepEqual(images.map((image) => imageColumn(image.range.tl)), [0, 1, 2, 3]);
});

function imageColumn(anchor: ExcelJS.ImageRange['tl']): number | undefined {
  const value = anchor as { readonly col?: number; readonly nativeCol?: number };
  return value.col ?? value.nativeCol;
}
