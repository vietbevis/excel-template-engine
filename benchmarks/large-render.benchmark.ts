import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { performance } from 'node:perf_hooks';
import ExcelJS from 'exceljs';
import { ExcelTemplateEngine } from '../src/application/engine/excel-template-engine.js';

const dir = await mkdtemp(join(tmpdir(), 'excel-template-engine-large-render-'));
const target = parseTarget();
const results: Partial<Record<BenchmarkTarget, BenchmarkResult>> = {};

if (target === 'all' || target === 'rows') {
  results.rows = await measure('render_100k_rows_each', async () => {
  const templatePath = join(dir, 'rows-100k.xlsx');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Rows').getCell('A1').value = '{{#each students}}{{code}}{{/each}}';
  await workbook.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    students: Array.from({ length: 100_000 }, (_value, index) => ({
      code: `SV${String(index + 1).padStart(6, '0')}`,
    })),
  }, {
    limits: {
      maxRows: 100_000,
      maxOperations: 110_000,
      maxTemplateBytes: 1024 * 1024,
    },
  });

  return { outputBytes: result.output.byteLength };
  });
}

if (target === 'all' || target === 'columns') {
  results.columns = await measure('render_5k_columns_each_col', async () => {
  const templatePath = join(dir, 'columns-5k.xlsx');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Columns').getCell('A1').value = '{{#each-col columns}}{{name}}{{/each-col}}';
  await workbook.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    columns: Array.from({ length: 5_000 }, (_value, index) => ({
      name: `C${index + 1}`,
    })),
  }, {
    limits: {
      maxColumns: 5_000,
      maxOperations: 6_000,
      maxTemplateBytes: 1024 * 1024,
    },
  });

  return { outputBytes: result.output.byteLength };
  });
}

if (target === 'all' || target === 'sheets') {
  results.sheets = await measure('render_50_worksheets', async () => {
  const templatePath = join(dir, 'worksheets-50.xlsx');
  const workbook = new ExcelJS.Workbook();
  for (let index = 0; index < 50; index += 1) {
    workbook.addWorksheet(`Sheet${index + 1}`).getCell('A1').value = '{{title}}';
  }
  await workbook.xlsx.writeFile(templatePath);

  const result = await new ExcelTemplateEngine().render(templatePath, {
    title: 'Report',
  }, {
    limits: {
      maxWorksheets: 50,
      maxOperations: 50,
      maxTemplateBytes: 1024 * 1024,
    },
  });

  return { outputBytes: result.output.byteLength };
  });
}

console.log(JSON.stringify(results, null, 2));

async function measure(name: string, run: () => Promise<{ readonly outputBytes: number }>): Promise<BenchmarkResult> {
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
    outputBytes: result.outputBytes,
    memory: {
      heapUsedDeltaMb: toMb(after.heapUsed - before.heapUsed),
      rssDeltaMb: toMb(after.rss - before.rss),
      heapUsedMb: toMb(after.heapUsed),
      rssMb: toMb(after.rss),
    },
  };
}

function toMb(bytes: number): number {
  return Math.round((bytes / 1024 / 1024) * 100) / 100;
}

interface BenchmarkResult {
  readonly name: string;
  readonly durationMs: number;
  readonly outputBytes: number;
  readonly memory: {
    readonly heapUsedDeltaMb: number;
    readonly rssDeltaMb: number;
    readonly heapUsedMb: number;
    readonly rssMb: number;
  };
}

type BenchmarkTarget = 'rows' | 'columns' | 'sheets';

function parseTarget(): BenchmarkTarget | 'all' {
  const rawTarget = process.argv
    .find((arg) => arg.startsWith('--target='))
    ?.slice('--target='.length);

  if (!rawTarget) {
    return 'all';
  }

  if (rawTarget === 'rows' || rawTarget === 'columns' || rawTarget === 'sheets') {
    return rawTarget;
  }

  throw new Error(`Unknown benchmark target: ${rawTarget}`);
}
