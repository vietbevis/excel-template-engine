import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import ExcelJS from 'exceljs';
import { ExcelTemplateEngine } from '../src/index.js';
import { DefaultAssetResolver } from '../src/infrastructure/assets/default-asset-resolver.js';
import { RenderPlanner } from '../src/application/planner/render-planner.js';
import { DefaultHelperRegistry } from '../src/core/evaluator/helper-registry.js';
import { EvaluationContext } from '../src/core/evaluator/evaluation-context.js';
import type { WorkbookAST } from '../src/core/ast/nodes.js';
import { LimitExceededError } from '../src/shared/errors/engine-error.js';

const onePixelPngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';

test('DefaultAssetResolver chặn absolute path mặc định và enforce maxBytes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-asset-limits-'));
  const imagePath = join(dir, 'avatar.png');
  await writeFile(imagePath, Buffer.from(onePixelPngBase64, 'base64'));

  await assert.rejects(
    () => new DefaultAssetResolver({ baseDir: dir }).resolve(imagePath),
    /Absolute image paths are disabled/,
  );

  await assert.rejects(
    () => new DefaultAssetResolver({ baseDir: dir, maxBytes: 4 }).resolve('./avatar.png'),
    LimitExceededError,
  );

  const asset = await new DefaultAssetResolver({ baseDir: dir }).resolve('./avatar.png');
  assert.equal(asset.extension, 'png');
});

test('DefaultAssetResolver chặn relative path thoát khỏi baseDir', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-asset-traversal-'));

  await assert.rejects(
    () => new DefaultAssetResolver({ baseDir: dir }).resolve('../avatar.png'),
    /escapes the configured baseDir/,
  );
});

test('RenderPlanner enforce workbook dimension và operation limits', async () => {
  const planner = new RenderPlanner(new DefaultHelperRegistry());
  const ast: WorkbookAST = {
    kind: 'WorkbookAST',
    sheets: [
      {
        kind: 'SheetAST',
        name: 'Sheet1',
        sheetId: '1',
        rows: [
          {
            kind: 'RowAST',
            rowNumber: 2,
            cells: [
              {
                kind: 'CellAST',
                address: { row: 2, column: 2 },
                rawValue: '{{name}}',
                nodes: [{
                  kind: 'PlaceholderNode',
                  id: 'placeholder_0',
                  path: 'name',
                  source: { start: 0, end: 8 },
                }],
              },
            ],
          },
        ],
      },
    ],
  };

  await assert.rejects(
    () => planner.createPlan(ast, EvaluationContext.root({ name: 'A' }), {
      limits: { maxRows: 1 },
    }),
    LimitExceededError,
  );

  await assert.rejects(
    () => planner.createPlan(ast, EvaluationContext.root({ name: 'A' }), {
      limits: { maxColumns: 1 },
    }),
    LimitExceededError,
  );

  await assert.rejects(
    () => planner.createPlan(ast, EvaluationContext.root({ name: 'A' }), {
      limits: { maxOperations: 0 },
    }),
    LimitExceededError,
  );
});

test('ExcelTemplateEngine enforce maxTemplateBytes trước khi load workbook', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-template-limit-'));
  const templatePath = join(dir, 'template.xlsx');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Report').getCell('A1').value = '{{name}}';
  await workbook.xlsx.writeFile(templatePath);

  const engine = new ExcelTemplateEngine();
  await assert.rejects(
    () => engine.render(templatePath, { name: 'A' }, {
      limits: { maxTemplateBytes: 1 },
    }),
    LimitExceededError,
  );
});
