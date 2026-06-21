import test from 'node:test';
import assert from 'node:assert/strict';
import { RenderPlanner } from '../src/application/planner/render-planner.js';
import { DefaultHelperRegistry } from '../src/core/evaluator/helper-registry.js';
import { EvaluationContext } from '../src/core/evaluator/evaluation-context.js';
import type { WorkbookAST } from '../src/core/ast/nodes.js';

test('RenderPlanner dùng ExpressionEvaluator cho placeholder default value', async () => {
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
            rowNumber: 1,
            cells: [
              {
                kind: 'CellAST',
                address: { row: 1, column: 1 },
                rawValue: '{{teacher.name ?? "Unknown"}}',
                nodes: [
                  {
                    kind: 'PlaceholderNode',
                    id: 'placeholder_0',
                    path: 'teacher.name ?? "Unknown"',
                    source: { start: 0, end: 29 },
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };

  const plan = await planner.createPlan(ast, EvaluationContext.root({ teacher: null }));

  assert.equal(plan.operations.length, 1);
  assert.deepEqual(plan.operations[0], {
    id: 'set_Sheet1_1_1',
    type: 'SetCellValue',
    sheetName: 'Sheet1',
    cell: { row: 1, column: 1 },
    value: 'Unknown',
  });
});
