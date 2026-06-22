import ExcelJS from 'exceljs';

export const THONG_KE_SOURCE_TEMPLATE_PATH = 'examples/templates/thong_ke_vuot_gio.xlsx';
export const THONG_KE_DATA_PATH = 'examples/data/thong-ke-vuot-gio-showcase-data.json';
export const THONG_KE_RENDER_TEMPLATE_PATH = 'examples/templates/thong-ke-vuot-gio-showcase-template.xlsx';
export const THONG_KE_RENDER_OUTPUT_PATH = 'examples/output/thong-ke-vuot-gio-showcase-output.xlsx';
export const THONG_KE_WORKSHEET_NAME = 'TK giảng dạy vượt giờ';

export async function createThongKeVuotGioShowcaseTemplate(): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(THONG_KE_SOURCE_TEMPLATE_PATH);

  const worksheet = workbook.getWorksheet(THONG_KE_WORKSHEET_NAME);
  if (!worksheet) {
    throw new Error(`Worksheet not found: ${THONG_KE_WORKSHEET_NAME}`);
  }

  const titleRow = snapshotRow(worksheet, 15, 7);
  const bodyRow = snapshotRow(worksheet, 19, 7);
  worksheet.getCell('A2').value = 'KHOA: {{metadata.facultyName}}';
  worksheet.getCell('D3').value = 'Hà nội, ngày {{metadata.reportDay}} tháng {{metadata.reportMonth}} năm {{metadata.reportYear}}';
  worksheet.getCell('A9').value = 'Họ và tên:  {{teacher.fullName}}';
  worksheet.getCell('D9').value = 'Ngày sinh: {{teacher.birthDate}}';
  worksheet.getCell('A10').value = 'Học hàm/ Học vị: {{teacher.academicTitle}}';

  unmergeRanges(worksheet, (range) => range.start.row >= 15);
  worksheet.spliceRows(15, Math.max(worksheet.rowCount - 14, 0));

  worksheet.getCell('A15').value = '{{#block sections}}';
  applyRowSnapshot(worksheet, 16, titleRow);
  worksheet.getCell('A16').value = '{{#each-col titleCells span=span reserve=7}}{{value}}{{/each-col}}';
  worksheet.getCell('A17').value = '{{#block rows}}';
  applyRowSnapshot(worksheet, 18, bodyRow);
  worksheet.getCell('A18').value = '{{#each-col cells span=span reserve=7}}{{value}}{{/each-col}}';
  worksheet.getCell('A19').value = '{{/block}}';
  worksheet.getCell('A20').value = '{{/block}}';

  for (let rowNumber = 21; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    worksheet.getRow(rowNumber).eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
    });
  }
  for (let rowNumber = worksheet.rowCount; rowNumber >= 21; rowNumber -= 1) {
    worksheet.spliceRows(rowNumber, 1);
  }

  await workbook.xlsx.writeFile(THONG_KE_RENDER_TEMPLATE_PATH);
  console.log(JSON.stringify({ output: THONG_KE_RENDER_TEMPLATE_PATH }, null, 2));
}

if (process.argv[1]?.endsWith('create-thong-ke-vuot-gio-showcase-template.js')) {
  await createThongKeVuotGioShowcaseTemplate();
}

interface A1Range {
  readonly start: { readonly row: number; readonly column: number };
  readonly end: { readonly row: number; readonly column: number };
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

function parseA1Cell(value: string): { readonly row: number; readonly column: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(value);
  if (!match) {
    throw new Error(`Invalid A1 cell: ${value}`);
  }

  return {
    column: columnNameToNumber(match[1]!),
    row: Number(match[2]),
  };
}

function columnNameToNumber(name: string): number {
  return [...name].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function clone<T>(value: T): T {
  return value === undefined ? value : JSON.parse(JSON.stringify(value)) as T;
}
