import { performance } from 'node:perf_hooks';
import { RenderPlanner } from '../src/application/planner/render-planner.js';
import { DefaultHelperRegistry } from '../src/core/evaluator/helper-registry.js';
import { EvaluationContext } from '../src/core/evaluator/evaluation-context.js';
import type { CellAST, TemplateNode, WorkbookAST } from '../src/core/ast/nodes.js';

const planner = new RenderPlanner(new DefaultHelperRegistry());

const rows100k = Array.from({ length: 100_000 }, (_value, index) => ({
  code: `SV${String(index + 1).padStart(6, '0')}`,
}));
const columns5k = Array.from({ length: 5_000 }, (_value, index) => ({
  name: `C${index + 1}`,
}));

const rowBenchmark = await measure('plan_100k_rows_each', async () => {
  const plan = await planner.createPlan(
    workbookWithCell('Rows', cell(1, 1, eachNode('students', placeholderNode('code')))),
    EvaluationContext.root({ students: rows100k }),
    { limits: { maxOperations: 110_000 } },
  );
  return { operations: plan.operations.length };
});

const columnBenchmark = await measure('plan_5k_columns_each_col', async () => {
  const plan = await planner.createPlan(
    workbookWithCell('Columns', cell(1, 1, eachColumnNode('columns', placeholderNode('name')))),
    EvaluationContext.root({ columns: columns5k }),
    { limits: { maxOperations: 6_000 } },
  );
  return { operations: plan.operations.length };
});

const sheetsBenchmark = await measure('plan_50_worksheets', async () => {
  const plan = await planner.createPlan(
    workbookWithSheets(50),
    EvaluationContext.root({ title: 'Report' }),
    { limits: { maxWorksheets: 50, maxOperations: 50 } },
  );
  return { operations: plan.operations.length };
});

console.log(JSON.stringify({
  rows100k: rowBenchmark,
  columns5k: columnBenchmark,
  worksheets50: sheetsBenchmark,
}, null, 2));

async function measure(
  name: string,
  run: () => Promise<{ readonly operations: number }>,
): Promise<BenchmarkResult> {
  if (global.gc) {
    global.gc();
  }

  const before = process.memoryUsage();
  const start = performance.now();
  const result = await run();
  const durationMs = performance.now() - start;
  const after = process.memoryUsage();

  return {
    name,
    durationMs: Math.round(durationMs),
    operations: result.operations,
    operationsPerSecond: Math.round(result.operations / (durationMs / 1000)),
    memory: {
      heapUsedDeltaMb: toMb(after.heapUsed - before.heapUsed),
      rssDeltaMb: toMb(after.rss - before.rss),
      heapUsedMb: toMb(after.heapUsed),
      rssMb: toMb(after.rss),
    },
  };
}

function workbookWithCell(sheetName: string, templateCell: CellAST): WorkbookAST {
  return {
    kind: 'WorkbookAST',
    sheets: [{
      kind: 'SheetAST',
      name: sheetName,
      sheetId: '1',
      rows: [{
        kind: 'RowAST',
        rowNumber: templateCell.address.row,
        cells: [templateCell],
      }],
    }],
  };
}

function workbookWithSheets(count: number): WorkbookAST {
  return {
    kind: 'WorkbookAST',
    sheets: Array.from({ length: count }, (_value, index) => ({
      kind: 'SheetAST',
      name: `Sheet${index + 1}`,
      sheetId: String(index + 1),
      rows: [{
        kind: 'RowAST',
        rowNumber: 1,
        cells: [cell(1, 1, placeholderNode('title'))],
      }],
    })),
  };
}

function cell(row: number, column: number, node: TemplateNode): CellAST {
  return {
    kind: 'CellAST',
    address: { row, column },
    nodes: [node],
    rawValue: '',
  };
}

function placeholderNode(path: string): TemplateNode {
  return {
    kind: 'PlaceholderNode',
    id: `placeholder_${path}`,
    path,
    source: { start: 0, end: path.length },
  };
}

function eachNode(path: string, child: TemplateNode): TemplateNode {
  return {
    kind: 'EachNode',
    id: `each_${path}`,
    path,
    children: [child],
    source: { start: 0, end: path.length },
  };
}

function eachColumnNode(path: string, child: TemplateNode): TemplateNode {
  return {
    kind: 'EachColumnNode',
    id: `each_col_${path}`,
    path,
    children: [child],
    source: { start: 0, end: path.length },
  };
}

function toMb(bytes: number): number {
  return Math.round((bytes / 1024 / 1024) * 100) / 100;
}

interface BenchmarkResult {
  readonly name: string;
  readonly durationMs: number;
  readonly operations: number;
  readonly operationsPerSecond: number;
  readonly memory: {
    readonly heapUsedDeltaMb: number;
    readonly rssDeltaMb: number;
    readonly heapUsedMb: number;
    readonly rssMb: number;
  };
}
