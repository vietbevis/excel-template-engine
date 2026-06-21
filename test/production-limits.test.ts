import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { DefaultAssetResolver } from '../src/infrastructure/assets/default-asset-resolver.js';
import { RenderPlanner } from '../src/application/planner/render-planner.js';
import { DefaultHelperRegistry } from '../src/core/evaluator/helper-registry.js';
import { EvaluationContext } from '../src/core/evaluator/evaluation-context.js';
import type { WorkbookAST } from '../src/core/ast/nodes.js';

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
    /Image exceeds maxBytes limit/,
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
    /exceeds maxRows limit/,
  );

  await assert.rejects(
    () => planner.createPlan(ast, EvaluationContext.root({ name: 'A' }), {
      limits: { maxColumns: 1 },
    }),
    /exceeds maxColumns limit/,
  );

  await assert.rejects(
    () => planner.createPlan(ast, EvaluationContext.root({ name: 'A' }), {
      limits: { maxOperations: 0 },
    }),
    /exceeds maxOperations limit/,
  );
});
