import ExcelJS from 'exceljs';
import { writeFile } from 'node:fs/promises';
import type {
  ColumnTreeBand,
  ColumnTreeColumnDefinition,
  ColumnTreeFormula,
  ColumnTreeStaticColumn,
} from '../src/index.js';
import {
  ColumnTreeCompiler,
  ExcelTemplateEngine,
  FormulaTemplateCompiler,
} from '../src/index.js';
import {
  OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME,
  OVERTIME_RENDER_OUTPUT_PATH,
  OVERTIME_RENDER_TEMPLATE_PATH,
  OVERTIME_WORKSHEET_NAME
} from './create-overtime-report-showcase-template.js';
import { overtimeReportShowcaseData } from './data/overtime-report-showcase-data.js';

const NUMBER_FORMAT = '#,##0.00';
const MONEY_FORMAT = '#,##0';

const templateLayout = await detectTemplateLayout(OVERTIME_RENDER_TEMPLATE_PATH);
const renderData = buildRenderData(overtimeReportShowcaseData, templateLayout);

const engine = new ExcelTemplateEngine();
const result = await engine.render(OVERTIME_RENDER_TEMPLATE_PATH, renderData);

await writeFile(OVERTIME_RENDER_OUTPUT_PATH, result.output);
console.log(JSON.stringify({
  output: OVERTIME_RENDER_OUTPUT_PATH,
  warnings: result.warnings,
}, null, 2));

function buildRenderData(data: OvertimeReportData, layout: OvertimeTemplateLayout): Record<string, unknown> {
  const compiledColumns = new ColumnTreeCompiler().compile(data.metricTree, {
    startColumn: layout.dynamicStartColumn,
  });
  const formulaCompiler = new FormulaTemplateCompiler(compiledColumns.columns, layout.staticFormulaColumns);
  const metricColumnCount = compiledColumns.columnCount;
  const sectionColSpan = layout.sectionHeaderStaticColumnSpan + metricColumnCount;
  const departmentBlock = data.blocks.departments;

  const departmentSheets = departmentBlock.items.map((department, departmentIndex) =>
    buildDepartmentSheetData(
      department,
      departmentIndex,
      data,
      compiledColumns.columns,
      formulaCompiler,
      layout.departmentSheet,
      metricColumnCount,
      sectionColSpan,
    ));

  let nextDepartmentHeaderRow = layout.firstDepartmentHeaderRow;
  const departmentTotalRows: number[] = [];
  const departments = departmentBlock.items.map((department, departmentIndex) => {
    const lecturerStartRow = nextDepartmentHeaderRow + 1;
    const lecturerEndRow = lecturerStartRow + department.lecturers.length - 1;
    const totalRow = lecturerEndRow + 1;
    departmentTotalRows.push(totalRow);
    nextDepartmentHeaderRow = totalRow + 1;
    const sheetName = department.name;
    const departmentSheetTotalRow = department.lecturers.length + layout.departmentSheet.firstLecturerRow;

    return {
      section: formula(sheetCell(sheetName, 'A', layout.departmentSheet.sectionHeaderRow)),
      name: department.name,
      sectionHeader: [{
        name: formula(sheetCell(sheetName, 'B', layout.departmentSheet.sectionHeaderRow)),
        colSpan: sectionColSpan,
      }],
      ...Object.fromEntries(departmentBlock.summary.staticTotals.map((total) => [
        total.field,
        formula(sheetCell(sheetName, staticTotalColumnName(layout.staticFormulaColumns, total.formula), departmentSheetTotalRow), NUMBER_FORMAT),
      ])),
      lecturers: department.lecturers.map((_lecturer, lecturerIndex) => ({
        displayIndex: formula(sheetCell(sheetName, 'A', layout.departmentSheet.firstLecturerRow + lecturerIndex)),
        name: formula(sheetCell(sheetName, 'B', layout.departmentSheet.firstLecturerRow + lecturerIndex)),
        salary: formula(sheetCell(sheetName, 'C', layout.departmentSheet.firstLecturerRow + lecturerIndex), MONEY_FORMAT),
        teachingNorm: formula(sheetCell(sheetName, 'D', layout.departmentSheet.firstLecturerRow + lecturerIndex), NUMBER_FORMAT),
        reduction: formula(sheetCell(sheetName, 'E', layout.departmentSheet.firstLecturerRow + lecturerIndex), NUMBER_FORMAT),
        researchIncomplete: formula(sheetCell(sheetName, 'F', layout.departmentSheet.firstLecturerRow + lecturerIndex), NUMBER_FORMAT),
        requiredHours: formula(sheetCell(sheetName, 'G', layout.departmentSheet.firstLecturerRow + lecturerIndex), NUMBER_FORMAT),
        metricValues: compiledColumns.columns.map((column) => ({
          value: column.fullKey === 'kyNhan.value'
            ? { value: null, wrapText: true }
            : formula(sheetCell(sheetName, column.columnName, layout.departmentSheet.firstLecturerRow + lecturerIndex), column.format?.numFmt),
        })),
      })),
      totals: compiledColumns.columns.map((column) => ({
        value: column.fullKey === 'kyNhan.value'
          ? { value: null, wrapText: true }
          : formula(sheetCell(sheetName, column.columnName, departmentSheetTotalRow), column.format?.numFmt),
      })),
    };
  });

  return {
    ...data,
    reportTitle: data.report.title,
    sheetTitle: data.report.summaryTitle,
    metricBands: compiledColumns.bands,
    metricGroups: compiledColumns.groups,
    metricColumns: compiledColumns.columns.map((column) => ({ name: column.name })),
    metricColumnCount,
    sectionColSpan,
    amountInWords: [{ text: `Bằng chữ: ${departmentBlock.amountInWords}`, colSpan: sectionColSpan }],
    _workbook: {
      worksheets: departmentSheets.map((departmentSheet, index) => ({
        sourceName: OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME,
        name: departmentSheet.sheetName,
        dataPath: `departmentSheets[${index}]`,
        deleteSource: true,
      })),
    },
    departmentSheets,
    ...Object.fromEntries(departmentBlock.summary.staticTotals.map((total) => [
      `grand${capitalize(total.field)}`,
      formula(sumCells(staticTotalColumnName(layout.staticFormulaColumns, total.formula), departmentTotalRows), NUMBER_FORMAT),
    ])),
    grandTotals: compiledColumns.columns.map((column) => ({
      value: column.fullKey === 'kyNhan.value'
        ? { value: null, wrapText: true }
        : formula(sumCells(column.columnName, departmentTotalRows), column.format?.numFmt),
    })),
    departments,
  };
}

function buildDepartmentSheetData(
  department: Department,
  departmentIndex: number,
  data: OvertimeReportData,
  columns: readonly ColumnTreeColumnDefinition[],
  formulaCompiler: FormulaTemplateCompiler,
  layout: DepartmentSheetTemplateLayout,
  metricColumnCount: number,
  sectionColSpan: number,
): Record<string, unknown> {
  const departmentBlock = data.blocks.departments;
  const lecturerStartRow = layout.firstLecturerRow;
  const lecturerEndRow = lecturerStartRow + department.lecturers.length - 1;
  const totalRow = lecturerEndRow + 1;
  const compiledColumns = new ColumnTreeCompiler().compile(data.metricTree, {
    startColumn: columns[0]?.columnName ? fromColumnName(columns[0].columnName) : 8,
  });

  return {
    ...data,
    sheetName: department.name,
    reportTitle: data.report.title,
    sheetTitle: department.title,
    section: department.section,
    name: department.name,
    metricBands: compiledColumns.bands,
    metricGroups: compiledColumns.groups,
    metricColumns: compiledColumns.columns.map((column) => ({ name: column.name })),
    metricColumnCount,
    sectionColSpan,
    sectionHeader: [{ name: department.name, colSpan: sectionColSpan }],
    amountInWords: [{ text: `Bằng chữ: ${departmentBlock.amountInWords}`, colSpan: sectionColSpan }],
    ...Object.fromEntries(departmentBlock.summary.staticTotals.map((total) => [
      total.field,
      formula(formulaCompiler.compile(total.formula, {
        row: totalRow,
        range: { startRow: lecturerStartRow, endRow: lecturerEndRow },
      }), NUMBER_FORMAT),
    ])),
    lecturers: department.lecturers.map((lecturer, lecturerIndex) => ({
      ...lecturer,
      salary: cellValue(lecturer.salary, MONEY_FORMAT),
      teachingNorm: cellValue(lecturer.teachingNorm, NUMBER_FORMAT),
      reduction: cellValue(lecturer.reduction, NUMBER_FORMAT),
      researchIncomplete: cellValue(lecturer.researchIncomplete, NUMBER_FORMAT),
      requiredHours: cellValue(lecturer.requiredHours, NUMBER_FORMAT),
      metricValues: toMetricValues(
        columns,
        formulaCompiler,
        lecturer.metrics,
        `blocks.departments.items[${departmentIndex}].lecturers[${lecturerIndex}].metrics`,
      ),
    })),
    totals: columns.map((column) => ({
      value: column.fullKey === 'kyNhan.value'
        ? { value: null, wrapText: true }
        : formula(formulaCompiler.compile(departmentBlock.summary.metricTotals.formula, {
          row: totalRow,
          currentRef: column.fullKey,
          range: { startRow: lecturerStartRow, endRow: lecturerEndRow },
        }), column.format?.numFmt),
    })),
  };
}

function toMetricValues(
  columns: readonly ColumnTreeColumnDefinition[],
  formulaCompiler: FormulaTemplateCompiler,
  metrics: MetricsByGroup,
  path: string,
): readonly { readonly value: unknown }[] {
  validateNoMissingMetricValues(columns, metrics, path);
  validateNoUnknownMetricValues(columns, metrics, path);

  return columns.map((column) => {
    if (column.formula) {
      return { value: formula(formulaCompiler.compile(column.formula), column.format?.numFmt) };
    }

    return { value: cellValue(readRequiredMetricValue(metrics, column, path), column.format?.numFmt) };
  });
}

function validateNoMissingMetricValues(
  columns: readonly ColumnTreeColumnDefinition[],
  metrics: MetricsByGroup,
  path: string,
): void {
  for (const column of columns) {
    if (!column.formula) {
      readRequiredMetricValue(metrics, column, path);
    }
  }
}

function validateNoUnknownMetricValues(
  columns: readonly ColumnTreeColumnDefinition[],
  metrics: MetricsByGroup,
  path: string,
): void {
  const expected = new Set(columns.filter((column) => !column.formula).map((column) => column.fullKey));
  for (const [groupKey, values] of Object.entries(metrics)) {
    if (typeof values !== 'object' || values === null || Array.isArray(values)) {
      throw new Error(`${path}.${groupKey} must be an object.`);
    }

    for (const columnKey of Object.keys(values)) {
      const fullKey = `${groupKey}.${columnKey}`;
      if (!expected.has(fullKey)) {
        throw new Error(`${path}.${fullKey} is not an input metric column. Check the key or remove formula/output-only data.`);
      }
    }
  }
}

function readRequiredMetricValue(
  metrics: MetricsByGroup,
  column: ColumnTreeColumnDefinition,
  path: string,
): unknown {
  const group = metrics[column.groupKey];
  if (typeof group !== 'object' || group === null || Array.isArray(group)) {
    throw new Error(`${path}.${column.groupKey} is required and must be an object.`);
  }

  if (!Object.prototype.hasOwnProperty.call(group, column.key)) {
    throw new Error(`${path}.${column.fullKey} is required.`);
  }

  return (group as Record<string, unknown>)[column.key];
}

function formula(value: string, numFmt?: string): { readonly value: { readonly formula: string }; readonly numFmt?: string } | { readonly formula: string } {
  if (!numFmt) {
    return { formula: value };
  }

  return {
    value: { formula: value },
    numFmt,
  };
}

function cellValue(value: unknown, numFmt?: string): unknown {
  if (!numFmt) {
    return value;
  }

  return {
    value,
    numFmt,
  };
}

function sumCells(columnName: string, rows: readonly number[]): string {
  if (rows.length === 0) {
    return '0';
  }

  return `SUM(${rows.map((row) => `${columnName}$${row}`).join(',')})`;
}

function sheetCell(sheetName: string, columnName: string, row: number): string {
  return `${quoteSheetName(sheetName)}!${columnName}$${row}`;
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_]+$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replace(/'/g, "''")}'`;
}

function staticTotalColumnName(columns: readonly ColumnTreeStaticColumn[], formula: ColumnTreeFormula): string {
  if (formula.type !== 'sumRange') {
    throw new Error(`Grand static total requires sumRange formula.`);
  }

  const key = formula.ref;
  const column = columns.find((item) => item.key === key);
  if (!column) {
    throw new Error(`Static total column not found: ${key}`);
  }

  return column.column;
}

function capitalize(value: string): string {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

async function detectTemplateLayout(templatePath: string): Promise<OvertimeTemplateLayout> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const worksheet = workbook.getWorksheet(OVERTIME_WORKSHEET_NAME);
  if (!worksheet) {
    throw new Error(`Worksheet not found: ${OVERTIME_WORKSHEET_NAME}`);
  }
  const departmentWorksheet = workbook.getWorksheet(OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME);
  if (!departmentWorksheet) {
    throw new Error(`Worksheet not found: ${OVERTIME_DEPARTMENT_TEMPLATE_SHEET_NAME}`);
  }

  const dynamicStartCell = findCell(worksheet, (value) => value.includes('{{#each-col metricBands'));
  const sectionHeaderCell = findCell(worksheet, (value) => value.includes('{{#each-col sectionHeader'));
  const departmentBlockStartCell = findCell(worksheet, (value) => value.trim() === '{{#block departments}}');
  const lecturerBodyRow = findCell(worksheet, (value) => value.trim() === '{{displayIndex}}').row;
  const departmentSectionHeaderRow = findCell(departmentWorksheet, (value) => value.trim() === '{{section}}').row;
  const departmentLecturerBlockStartRow = findCell(departmentWorksheet, (value) => value.trim() === '{{#block lecturers}}').row;

  return {
    dynamicStartColumn: dynamicStartCell.column,
    sectionHeaderStaticColumnSpan: dynamicStartCell.column - sectionHeaderCell.column,
    firstDepartmentHeaderRow: departmentBlockStartCell.row,
    staticFormulaColumns: detectStaticFormulaColumns(worksheet, lecturerBodyRow, dynamicStartCell.column),
    departmentSheet: {
      sectionHeaderRow: departmentSectionHeaderRow,
      firstLecturerRow: departmentLecturerBlockStartRow,
    },
  };
}

function detectStaticFormulaColumns(
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  dynamicStartColumn: number,
): readonly ColumnTreeStaticColumn[] {
  const columns: ColumnTreeStaticColumn[] = [];
  const row = worksheet.getRow(rowNumber);
  for (let columnNumber = 1; columnNumber < dynamicStartColumn; columnNumber += 1) {
    const value = row.getCell(columnNumber).value;
    if (typeof value !== 'string') {
      continue;
    }

    const match = /^\{\{([A-Za-z_][A-Za-z0-9_]*)}}$/.exec(value.trim());
    if (match?.[1]) {
      columns.push({ key: match[1], column: toColumnName(columnNumber) });
    }
  }

  return columns;
}

function findCell(
  worksheet: ExcelJS.Worksheet,
  predicate: (value: string) => boolean,
): { readonly row: number; readonly column: number } {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const value = row.getCell(columnNumber).value;
      if (typeof value === 'string' && predicate(value)) {
        return { row: rowNumber, column: columnNumber };
      }
    }
  }

  throw new Error('Cannot detect required overtime template cell.');
}

function toColumnName(columnNumber: number): string {
  let current = columnNumber;
  let name = '';
  while (current > 0) {
    current -= 1;
    name = String.fromCharCode(65 + (current % 26)) + name;
    current = Math.floor(current / 26);
  }

  return name;
}

function fromColumnName(columnName: string): number {
  return columnName.toUpperCase().split('').reduce((total, char) => {
    return total * 26 + char.charCodeAt(0) - 64;
  }, 0);
}

export interface OvertimeReportData {
  readonly schemaVersion: 1;
  readonly report: {
    readonly title: string;
    readonly summaryTitle: string;
  };
  readonly metricTree: readonly ColumnTreeBand[];
  readonly blocks: {
    readonly departments: {
      readonly amountInWords: string;
      readonly summary: {
        readonly staticTotals: ReadonlyArray<{
          readonly field: string;
          readonly formula: ColumnTreeFormula;
        }>;
        readonly metricTotals: {
          readonly formula: ColumnTreeFormula;
        };
      };
      readonly items: readonly Department[];
    };
  };
}

export interface OvertimeTemplateLayout {
  readonly dynamicStartColumn: number;
  readonly sectionHeaderStaticColumnSpan: number;
  readonly firstDepartmentHeaderRow: number;
  readonly staticFormulaColumns: readonly ColumnTreeStaticColumn[];
  readonly departmentSheet: DepartmentSheetTemplateLayout;
}

export interface DepartmentSheetTemplateLayout {
  readonly sectionHeaderRow: number;
  readonly firstLecturerRow: number;
}

export type MetricsByGroup = Record<string, Record<string, unknown>>;

export interface Department {
  readonly section: string;
  readonly name: string;
  readonly title: string;
  readonly lecturers: readonly Lecturer[];
}

export interface Lecturer {
  readonly displayIndex: number;
  readonly name: string;
  readonly salary: number;
  readonly teachingNorm: number;
  readonly reduction: number;
  readonly researchIncomplete: number;
  readonly requiredHours: number;
  readonly metrics: MetricsByGroup;
}
