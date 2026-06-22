import ExcelJS from 'exceljs';
import { overtimeReportShowcaseData } from './data/overtime-report-showcase-data.js';

export const OVERTIME_DATA_PATH = 'examples/data/overtime-report-showcase-data.json';
export const OVERTIME_SOURCE_TEMPLATE_PATH = 'examples/templates/template.xlsx';
export const OVERTIME_RENDER_TEMPLATE_PATH = 'examples/templates/overtime-report-showcase-template.xlsx';
export const OVERTIME_RENDER_OUTPUT_PATH = 'examples/output/overtime-report-showcase-output.xlsx';
export const OVERTIME_WORKSHEET_NAME = 'OvertimeReport';
export const OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME = 'DepartmentTemplate';

export async function createOvertimeReportShowcaseTemplate(): Promise<void> {
  const templateData = overtimeReportShowcaseData;
  const metricColumnCount = countMetricColumns(templateData.metricTree);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(OVERTIME_SOURCE_TEMPLATE_PATH);

  const worksheet = workbook.getWorksheet('Sheet1');
  if (!worksheet) {
    throw new Error(`Worksheet "Sheet1" was not found in ${OVERTIME_SOURCE_TEMPLATE_PATH}.`);
  }
  worksheet.name = OVERTIME_WORKSHEET_NAME;

  const dynamicStartColumn = detectDynamicStartColumn(worksheet);
  const maxColumn = dynamicStartColumn + metricColumnCount - 1;

  prepareOvertimeWorksheetBase(worksheet, maxColumn);
  const departmentTemplate = cloneWorksheet(workbook, worksheet, OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME);
  prepareOvertimeWorksheetBase(departmentTemplate, maxColumn);

  configureSummaryWorksheet(worksheet, dynamicStartColumn, maxColumn);
  configureDepartmentTemplateWorksheet(departmentTemplate, dynamicStartColumn, maxColumn);

  await workbook.xlsx.writeFile(OVERTIME_RENDER_TEMPLATE_PATH);
  console.log(JSON.stringify({ output: OVERTIME_RENDER_TEMPLATE_PATH }, null, 2));
}

function prepareOvertimeWorksheetBase(worksheet: ExcelJS.Worksheet, maxColumn: number): void {
  ensureStyledColumns(worksheet, maxColumn);
  mergeFullWidthRow(worksheet, 1, maxColumn);
  mergeFullWidthRow(worksheet, 2, maxColumn);
}

function configureSummaryWorksheet(
  worksheet: ExcelJS.Worksheet,
  dynamicStartColumn: number,
  maxColumn: number,
): void {
  const dynamicEndColumn = maxColumn;
  const totalRowStyle = snapshotRow(worksheet, 19, maxColumn);
  const wordsRowStyle = snapshotRow(worksheet, 20, maxColumn);

  unmergeRanges(worksheet, (range) => range.start.row >= 4 && range.end.column >= dynamicStartColumn);
  unmergeRanges(worksheet, (range) => range.start.row >= 7);

  for (let rowNumber = 7; rowNumber <= 15; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = null;
    }
  }
  for (let rowNumber = 4; rowNumber <= 6; rowNumber += 1) {
    for (let columnNumber = dynamicStartColumn; columnNumber <= dynamicEndColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = null;
    }
  }

  for (let rowNumber = 16; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
    });
  }
  worksheet.spliceRows(16, Math.max(worksheet.rowCount - 15, 0));

  worksheet.getCell('A1').value = '{{reportTitle}}';
  worksheet.getCell('A2').value = '{{sheetTitle}}';

  worksheet.getCell(4, dynamicStartColumn).value = '{{#each-col metricBands span=colCount rowspan=rowSpan reserve=metricColumnCount}}{{name}}{{/each-col}}';
  worksheet.getCell(5, dynamicStartColumn).value = '{{#each-col metricGroups span=colCount render=renderHeader reserve=metricColumnCount}}{{name}}{{/each-col}}';
  worksheet.getCell(6, dynamicStartColumn).value = '{{#each-col metricColumns reserve=metricColumnCount}}{{name}}{{/each-col}}';

  worksheet.getCell('A7').value = '{{#block departments}}';
  worksheet.getCell('A8').value = '{{section}}';
  worksheet.getCell('B8').value = '{{#each-col sectionHeader span=colSpan reserve=sectionColSpan}}{{name}}{{/each-col}}';

  worksheet.getCell('A9').value = '{{#block lecturers}}';
  worksheet.getCell('A10').value = '{{displayIndex}}';
  worksheet.getCell('B10').value = '{{name}}';
  worksheet.getCell('C10').value = '{{salary}}';
  worksheet.getCell('D10').value = '{{teachingNorm}}';
  worksheet.getCell('E10').value = '{{reduction}}';
  worksheet.getCell('F10').value = '{{researchIncomplete}}';
  worksheet.getCell('G10').value = '{{requiredHours}}';
  worksheet.getCell(10, dynamicStartColumn).value = '{{#each-col metricValues reserve=metricColumnCount}}{{value}}{{/each-col}}';
  worksheet.getCell('A11').value = '{{/block}}';

  applyRowSnapshot(worksheet, 12, totalRowStyle);
  worksheet.getCell('A12').value = 'Tổng cộng';
  worksheet.getCell('D12').value = '{{totalTeachingNorm}}';
  worksheet.getCell('E12').value = '{{totalReduction}}';
  worksheet.getCell('F12').value = '{{totalResearchIncomplete}}';
  worksheet.getCell('G12').value = '{{totalRequiredHours}}';
  worksheet.getCell(12, dynamicStartColumn).value = '{{#each-col totals reserve=metricColumnCount}}{{value}}{{/each-col}}';

  worksheet.getCell('A13').value = '{{/block}}';

  applyRowSnapshot(worksheet, 14, totalRowStyle);
  worksheet.getCell('A14').value = 'Tổng cộng toàn bộ';
  worksheet.getCell('D14').value = '{{grandTotalTeachingNorm}}';
  worksheet.getCell('E14').value = '{{grandTotalReduction}}';
  worksheet.getCell('F14').value = '{{grandTotalResearchIncomplete}}';
  worksheet.getCell('G14').value = '{{grandTotalRequiredHours}}';
  worksheet.getCell(14, dynamicStartColumn).value = '{{#each-col grandTotals reserve=metricColumnCount}}{{value}}{{/each-col}}';

  applyRowSnapshot(worksheet, 15, wordsRowStyle);
  worksheet.getCell('B15').value = '{{#each-col amountInWords span=colSpan reserve=sectionColSpan}}{{text}}{{/each-col}}';
}

function configureDepartmentTemplateWorksheet(
  worksheet: ExcelJS.Worksheet,
  dynamicStartColumn: number,
  maxColumn: number,
): void {
  const dynamicEndColumn = maxColumn;
  const totalRowStyle = snapshotRow(worksheet, 19, maxColumn);
  const wordsRowStyle = snapshotRow(worksheet, 20, maxColumn);

  unmergeRanges(worksheet, (range) => range.start.row >= 4 && range.end.column >= dynamicStartColumn);
  unmergeRanges(worksheet, (range) => range.start.row >= 7);

  for (let rowNumber = 7; rowNumber <= 15; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = null;
    }
  }
  for (let rowNumber = 4; rowNumber <= 6; rowNumber += 1) {
    for (let columnNumber = dynamicStartColumn; columnNumber <= dynamicEndColumn; columnNumber += 1) {
      worksheet.getCell(rowNumber, columnNumber).value = null;
    }
  }

  for (let rowNumber = 14; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
    });
  }
  worksheet.spliceRows(14, Math.max(worksheet.rowCount - 13, 0));

  worksheet.getCell('A1').value = '{{reportTitle}}';
  worksheet.getCell('A2').value = '{{sheetTitle}}';

  worksheet.getCell(4, dynamicStartColumn).value = '{{#each-col metricBands span=colCount rowspan=rowSpan reserve=metricColumnCount}}{{name}}{{/each-col}}';
  worksheet.getCell(5, dynamicStartColumn).value = '{{#each-col metricGroups span=colCount render=renderHeader reserve=metricColumnCount}}{{name}}{{/each-col}}';
  worksheet.getCell(6, dynamicStartColumn).value = '{{#each-col metricColumns reserve=metricColumnCount}}{{name}}{{/each-col}}';

  worksheet.getCell('A7').value = '{{section}}';
  worksheet.getCell('B7').value = '{{#each-col sectionHeader span=colSpan reserve=sectionColSpan}}{{name}}{{/each-col}}';
  worksheet.getRow(7).hidden = true;

  worksheet.getCell('A8').value = '{{#block lecturers}}';
  worksheet.getCell('A9').value = '{{displayIndex}}';
  worksheet.getCell('B9').value = '{{name}}';
  worksheet.getCell('C9').value = '{{salary}}';
  worksheet.getCell('D9').value = '{{teachingNorm}}';
  worksheet.getCell('E9').value = '{{reduction}}';
  worksheet.getCell('F9').value = '{{researchIncomplete}}';
  worksheet.getCell('G9').value = '{{requiredHours}}';
  worksheet.getCell(9, dynamicStartColumn).value = '{{#each-col metricValues reserve=metricColumnCount}}{{value}}{{/each-col}}';
  worksheet.getCell('A10').value = '{{/block}}';

  applyRowSnapshot(worksheet, 11, totalRowStyle);
  worksheet.getCell('A11').value = 'Tổng cộng';
  worksheet.getCell('D11').value = '{{totalTeachingNorm}}';
  worksheet.getCell('E11').value = '{{totalReduction}}';
  worksheet.getCell('F11').value = '{{totalResearchIncomplete}}';
  worksheet.getCell('G11').value = '{{totalRequiredHours}}';
  worksheet.getCell(11, dynamicStartColumn).value = '{{#each-col totals reserve=metricColumnCount}}{{value}}{{/each-col}}';

  applyRowSnapshot(worksheet, 12, wordsRowStyle);
  worksheet.getCell('B12').value = '{{#each-col amountInWords span=colSpan reserve=sectionColSpan}}{{text}}{{/each-col}}';
}

if (process.argv[1]?.endsWith('create-overtime-report-showcase-template.js')) {
  await createOvertimeReportShowcaseTemplate();
}

interface A1Range {
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
}

interface OvertimeTemplateData {
  readonly metricTree: readonly MetricTreeNode[];
}

interface MetricTreeNode {
  readonly columns?: readonly unknown[];
  readonly groups?: readonly MetricTreeNode[];
  readonly derive?: {
    readonly type: 'sumSameKey';
    readonly columns?: readonly string[];
    readonly total?: unknown;
  };
}

interface RowSnapshot {
  readonly height?: number;
  readonly cells: readonly Partial<ExcelJS.Cell>[];
}

function snapshotRow(worksheet: ExcelJS.Worksheet, rowNumber: number, maxColumn: number): RowSnapshot {
  const row = worksheet.getRow(rowNumber);
  const cells: Partial<ExcelJS.Cell>[] = [];
  for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
    const cell = row.getCell(columnNumber);
    cells.push({
      style: clone(cell.style),
      numFmt: cell.numFmt,
      alignment: clone(cell.alignment),
      border: clone(cell.border),
      fill: clone(cell.fill),
      font: clone(cell.font),
      protection: clone(cell.protection),
    });
  }

  return {
    ...(row.height ? { height: row.height } : {}),
    cells,
  };
}

function countMetricColumns(nodes: readonly MetricTreeNode[]): number {
  return nodes.reduce((total, node) => {
    if (node.groups) {
      return total + countMetricColumns(node.groups);
    }

    if (node.derive?.type === 'sumSameKey') {
      return total + (node.derive.columns?.length ?? 0) + (node.derive.total ? 1 : 0);
    }

    return total + (node.columns?.length ?? 0);
  }, 0);
}

function applyRowSnapshot(worksheet: ExcelJS.Worksheet, rowNumber: number, snapshot: RowSnapshot): void {
  const row = worksheet.getRow(rowNumber);
  if (snapshot.height) {
    row.height = snapshot.height;
  }

  for (let index = 0; index < snapshot.cells.length; index += 1) {
    const source = snapshot.cells[index]!;
    const cell = row.getCell(index + 1);
    if (source.style) {
      cell.style = clone(source.style) as Partial<ExcelJS.Style>;
    }
  }
}

function cloneWorksheet(
  workbook: ExcelJS.Workbook,
  source: ExcelJS.Worksheet,
  targetName: string,
): ExcelJS.Worksheet {
  const target = workbook.addWorksheet(targetName, {
    properties: clone(source.properties),
    pageSetup: clone(source.pageSetup),
    views: clone(source.views),
    state: source.state,
  });

  for (let columnNumber = 1; columnNumber <= source.columnCount; columnNumber += 1) {
    const sourceColumn = source.getColumn(columnNumber);
    const targetColumn = target.getColumn(columnNumber);
    if (sourceColumn.width !== undefined) {
      targetColumn.width = sourceColumn.width;
    }
    if (sourceColumn.hidden !== undefined) {
      targetColumn.hidden = sourceColumn.hidden;
    }
    if (sourceColumn.outlineLevel !== undefined) {
      targetColumn.outlineLevel = sourceColumn.outlineLevel;
    }
    if (sourceColumn.style !== undefined) {
      targetColumn.style = clone(sourceColumn.style) as Partial<ExcelJS.Style>;
    }
  }

  source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    if (sourceRow.height !== undefined) {
      targetRow.height = sourceRow.height;
    }
    if (sourceRow.hidden !== undefined) {
      targetRow.hidden = sourceRow.hidden;
    }
    if (sourceRow.outlineLevel !== undefined) {
      targetRow.outlineLevel = sourceRow.outlineLevel;
    }

    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, columnNumber) => {
      const targetCell = targetRow.getCell(columnNumber);
      targetCell.value = clone(sourceCell.value);
      targetCell.style = clone(sourceCell.style) as Partial<ExcelJS.Style>;
      targetCell.dataValidation = clone(sourceCell.dataValidation);
      if (hasMeaningfulNote(sourceCell.note)) {
        targetCell.note = clone(sourceCell.note);
      }
    });
  });

  const mergeModel = source as unknown as { _merges?: Record<string, { readonly range: string }> };
  for (const merge of Object.values(mergeModel._merges ?? {})) {
    target.mergeCells(merge.range);
  }

  return target;
}

function hasMeaningfulNote(note: ExcelJS.Cell['note']): note is NonNullable<ExcelJS.Cell['note']> {
  if (note === undefined || note === null || note === '') {
    return false;
  }

  if (typeof note === 'string') {
    return note.trim() !== '';
  }

  if (typeof note === 'object' && 'texts' in note && Array.isArray(note.texts)) {
    return note.texts.some((text) => typeof text.text === 'string' && text.text.trim() !== '');
  }

  return true;
}

function detectDynamicStartColumn(worksheet: ExcelJS.Worksheet): number {
  const headerStartRow = 4;
  const headerEndRow = 6;
  const mergeModel = worksheet as unknown as { _merges?: Record<string, { range: string }> };
  const verticalHeaderColumns = Object.values(mergeModel._merges ?? {})
    .map((merge) => parseA1Range(merge.range))
    .filter((range) => range.start.row === headerStartRow && range.end.row === headerEndRow)
    .filter((range) => range.start.column === range.end.column)
    .map((range) => range.start.column)
    .sort((left, right) => left - right);

  let staticHeaderEndColumn = 0;
  for (const column of verticalHeaderColumns) {
    if (column !== staticHeaderEndColumn + 1) {
      break;
    }

    staticHeaderEndColumn = column;
  }

  if (staticHeaderEndColumn <= 0) {
    throw new Error('Cannot detect dynamic header start column from source template merges.');
  }

  return staticHeaderEndColumn + 1;
}

function ensureStyledColumns(worksheet: ExcelJS.Worksheet, maxColumn: number): void {
  const sourceMaxColumn = worksheet.columnCount;
  if (sourceMaxColumn >= maxColumn) {
    return;
  }

  for (let columnNumber = sourceMaxColumn + 1; columnNumber <= maxColumn; columnNumber += 1) {
    const sourceColumnNumber = columnNumber - 1;
    const sourceWidth = worksheet.getColumn(sourceColumnNumber).width;
    if (sourceWidth !== undefined) {
      worksheet.getColumn(columnNumber).width = sourceWidth;
    }
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const source = worksheet.getCell(rowNumber, sourceColumnNumber);
      const target = worksheet.getCell(rowNumber, columnNumber);
      target.style = clone(source.style) as Partial<ExcelJS.Style>;
    }
  }
}

function mergeFullWidthRow(worksheet: ExcelJS.Worksheet, rowNumber: number, maxColumn: number): void {
  unmergeRanges(worksheet, (range) => range.start.row <= rowNumber && range.end.row >= rowNumber);
  worksheet.mergeCells(rowNumber, 1, rowNumber, maxColumn);
}

function unmergeRanges(worksheet: ExcelJS.Worksheet, predicate: (range: A1Range) => boolean): void {
  const mergeModel = worksheet as unknown as { _merges?: Record<string, { range: string }> };
  const ranges = Object.values(mergeModel._merges ?? {}).map((merge) => merge.range);
  for (const rangeText of ranges) {
    const range = parseA1Range(rangeText);
    if (predicate(range)) {
      worksheet.unMergeCells(rangeText);
    }
  }
}

function parseA1Range(value: string): A1Range {
  const [start, end = start] = value.split(':');
  return {
    start: parseA1Cell(start!),
    end: parseA1Cell(end!),
  };
}

function parseA1Cell(value: string): { row: number; column: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(value);
  if (!match) {
    throw new Error(`Invalid A1 cell: ${value}`);
  }

  let column = 0;
  for (const character of match[1]!) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }

  return {
    row: Number(match[2]),
    column,
  };
}

function columnNameToNumber(columnName: string): number {
  let number = 0;
  for (const character of columnName.toUpperCase()) {
    number = number * 26 + character.charCodeAt(0) - 64;
  }

  return number;
}

function clone<T>(value: T): T {
  return value ? JSON.parse(JSON.stringify(value)) as T : value;
}
