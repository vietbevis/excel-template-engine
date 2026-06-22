import { readFile, writeFile } from 'node:fs/promises';
import { ExcelTemplateEngine } from '../src/index.js';
import {
  createThongKeVuotGioShowcaseTemplate,
  THONG_KE_DATA_PATH,
  THONG_KE_RENDER_OUTPUT_PATH,
  THONG_KE_RENDER_TEMPLATE_PATH,
} from './create-thong-ke-vuot-gio-showcase-template.js';

const COLUMN_KEYS = ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7'] as const;

await createThongKeVuotGioShowcaseTemplate();

const renderData = JSON.parse(await readFile(THONG_KE_DATA_PATH, 'utf8')) as ThongKeVuotGioRenderData;
const compiledData = compileRenderData(renderData);
const engine = new ExcelTemplateEngine();
const result = await engine.render(
  THONG_KE_RENDER_TEMPLATE_PATH,
  compiledData as unknown as Record<string, unknown>,
  { preserveFormulas: true },
);

await writeFile(THONG_KE_RENDER_OUTPUT_PATH, result.output);
console.log(JSON.stringify({
  output: THONG_KE_RENDER_OUTPUT_PATH,
  warnings: result.warnings,
}, null, 2));

function compileRenderData(data: ThongKeVuotGioRenderData): ThongKeVuotGioCompiledData {
  const state: CompileState = {
    nextRow: 15,
    refs: new Map(),
  };

  return {
    metadata: data.metadata,
    teacher: data.teacher,
    sections: compileSections(data.sections, state),
  };
}

function compileSections(sections: readonly ReportSection[], state: CompileState): readonly RenderSection[] {
  const compiled: RenderSection[] = [];
  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index]!;
    const nextSection = sections[index + 1];
    if (section.blocks.length === 0 && nextSection) {
      state.nextRow += 1;
      compiled.push({
        titleCells: [{ value: section.title, span: 7 }],
        rows: [
          fullWidthRow(nextSection.title, state, nextSection.id),
          ...nextSection.blocks.flatMap((block) => compileBlock(nextSection, block, state)),
        ],
      });
      index += 1;
      continue;
    }

    state.nextRow += 1;
    compiled.push({
      titleCells: [{ value: section.title, span: 7 }],
      rows: section.blocks.flatMap((block) => compileBlock(section, block, state)),
    });
  }

  return compiled;
}

function compileBlock(section: ReportSection, block: ReportBlock, state: CompileState): readonly RenderRow[] {
  if (block.type === 'heading') {
    return [fullWidthRow(block.text, state, `${section.id}.${block.id}`)];
  }

  if (block.type === 'summary') {
    return [summaryRow(block, state, `${section.id}.${block.id}`)];
  }

  if (block.type === 'row') {
    return [valuesRow(Object.values(block.values), state, `${section.id}.${block.id}`)];
  }

  const rows: RenderRow[] = [valuesRow(block.columns.map((column) => column.name), state, `${section.id}.${block.id}.header`)];
  for (const group of block.groups) {
    const groupRef = `${section.id}.${group.id}`;
    if (group.title) {
      rows.push(fullWidthRow(group.title, state, groupRef));
    }

    const itemRows: number[] = [];
    for (const item of group.items) {
      const result = tableItemRow(block.columns, item.values, state, `${groupRef}.${item.id}`);
      itemRows.push(result.rowNumber);
      rows.push(result.row);
    }

    if (group.total) {
      rows.push(summaryRow(group.total, state, `${groupRef}.total`, { itemRows }));
    }
  }

  rows.push(...block.summaryRows.map((row) => summaryRow(row, state, `${section.id}.${block.id}.${row.id}`)));
  return rows;
}

function fullWidthRow(value: RenderValue, state: CompileState, refId: string): RenderRow {
  const rowNumber = reserveRow(state, refId);
  return { cells: [{ value: resolveValue(value, state, { rowNumber }), span: 7 }] };
}

function valuesRow(
  values: ReadonlyArray<RenderValue | undefined>,
  state: CompileState,
  refId: string,
  context: FormulaContext = {},
): RenderRow {
  const rowNumber = reserveRow(state, refId);
  return {
    cells: normalizeValues(values).map((value) => ({ value: resolveValue(value, state, { ...context, rowNumber }), span: 1 })),
  };
}

function tableItemRow(
  columns: readonly ReportColumn[],
  item: ReportValues,
  state: CompileState,
  refId: string,
): { readonly row: RenderRow; readonly rowNumber: number } {
  const rowNumber = state.nextRow;
  return {
    row: valuesRow(columns.map((column) => item[column.key] ?? ''), state, refId),
    rowNumber,
  };
}

function summaryRow(
  summary: ReportSummary,
  state: CompileState,
  refId: string,
  context: FormulaContext = {},
): RenderRow {
  const values = normalizeValues(toColumnValues(summary.values));
  const firstFiveAreLabel = values.slice(0, 5).every((value) => value === summary.label);
  const rowNumber = reserveRow(state, refId);
  if (!firstFiveAreLabel) {
    return {
      cells: values.map((value) => ({ value: resolveValue(value, state, { ...context, rowNumber }), span: 1 })),
    };
  }

  return {
    cells: [
      { value: summary.label, span: 5 },
      { value: resolveValue(values[5] ?? '', state, { ...context, rowNumber }), span: 1 },
      { value: resolveValue(values[6] ?? '', state, { ...context, rowNumber }), span: 1 },
    ],
  };
}

function reserveRow(state: CompileState, refId: string): number {
  const rowNumber = state.nextRow;
  state.nextRow += 1;
  for (const key of COLUMN_KEYS) {
    state.refs.set(`${refId}.${key}`, `${columnName(key)}${rowNumber}`);
  }

  return rowNumber;
}

function resolveValue(value: RenderValue, state: CompileState, context: FormulaContext): RenderValue {
  if (!isFormulaContainer(value)) {
    return value;
  }

  return { formula: compileFormula(value.formula, state, context) };
}

function compileFormula(formula: FormulaSpec, state: CompileState, context: FormulaContext): string {
  if (typeof formula === 'string') {
    return formula;
  }

  if (formula.type === 'sumItems') {
    const rows = context.itemRows ?? [];
    const column = columnName(formula.column);
    const expression = rows.length === 0
      ? '0'
      : rows.length === 1
        ? `${column}${rows[0]}`
        : `SUM(${column}${rows[0]}:${column}${rows[rows.length - 1]})`;
    return formula.round ? `${formula.round.fn}(${expression},${formula.round.digits})` : expression;
  }

  return formula.template.replace(/\{([A-Za-z0-9_.]+)\}/g, (match, ref: string) => {
    return state.refs.get(ref) ?? match;
  });
}

function isFormulaContainer(value: RenderValue): value is { readonly formula: FormulaSpec } {
  return typeof value === 'object'
    && value !== null
    && 'formula' in value;
}

function toColumnValues(values: ReportValues): readonly RenderValue[] {
  return COLUMN_KEYS.map((key) => values[key] ?? '');
}

function normalizeValues(values: ReadonlyArray<RenderValue | undefined>): readonly RenderValue[] {
  return Array.from({ length: 7 }, (_, index) => values[index] ?? '');
}

function columnName(key: string): string {
  return String.fromCharCode(64 + Number(key.slice(1)));
}


interface ThongKeVuotGioRenderData {
  readonly schemaVersion: 1;
  readonly metadata: {
    readonly facultyName: string;
    readonly reportDay: string;
    readonly reportMonth: string;
    readonly reportYear: number;
  };
  readonly teacher: TeacherData;
  readonly sections: readonly ReportSection[];
}

interface TeacherData {
  readonly fullName: string;
  readonly birthDate: string;
  readonly academicTitle: string;
}

interface ReportSection {
  readonly id: string;
  readonly title: string;
  readonly blocks: readonly ReportBlock[];
}

type ReportBlock =
  | ReportHeadingBlock
  | ReportRowBlock
  | ReportTableBlock
  | ReportSummaryBlock;

interface ReportHeadingBlock {
  readonly type: 'heading';
  readonly id: string;
  readonly text: string;
}

interface ReportRowBlock {
  readonly type: 'row';
  readonly id: string;
  readonly values: ReportValues;
}

interface ReportSummaryBlock extends ReportSummary {
  readonly type: 'summary';
}

interface ReportTableBlock {
  readonly id: string;
  readonly type: 'table';
  readonly columns: readonly ReportColumn[];
  readonly groups: readonly ReportTableGroup[];
  readonly summaryRows: readonly ReportSummary[];
}

interface ReportColumn {
  readonly key: string;
  readonly name: string;
}

interface ReportTableGroup {
  readonly id: string;
  readonly title: string;
  readonly items: readonly ReportTableItem[];
  readonly total?: ReportSummary;
}

interface ReportTableItem {
  readonly id: string;
  readonly sourceRow?: number;
  readonly values: ReportValues;
}

interface ReportSummary {
  readonly id: string;
  readonly label: string;
  readonly values: ReportValues;
}

interface ReportValues {
  readonly [key: string]: RenderValue | undefined;
}

interface ThongKeVuotGioCompiledData {
  readonly metadata: ThongKeVuotGioRenderData['metadata'];
  readonly teacher: TeacherData;
  readonly sections: readonly RenderSection[];
}

interface RenderSection {
  readonly titleCells: readonly RenderCell[];
  readonly rows: readonly RenderRow[];
}

interface RenderRow {
  readonly cells: readonly RenderCell[];
}

interface RenderCell {
  readonly value: RenderValue;
  readonly span: number;
}

type RenderValue =
  | string
  | number
  | boolean
  | { readonly formula: FormulaSpec }
  | RichTextValue;

type FormulaSpec =
  | string
  | {
    readonly type: 'excel';
    readonly template: string;
  }
  | {
    readonly type: 'sumItems';
    readonly column: string;
    readonly round?: {
      readonly fn: 'TRUNC' | 'ROUND';
      readonly digits: number;
    };
  };

interface CompileState {
  nextRow: number;
  refs: Map<string, string>;
}

interface FormulaContext {
  readonly rowNumber?: number;
  readonly itemRows?: readonly number[];
}

interface RichTextValue {
  readonly richText: ReadonlyArray<{
    readonly text?: string;
    readonly font?: unknown;
  }>;
}
