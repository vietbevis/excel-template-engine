import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
workbook.creator = 'excel-template-engine';
workbook.created = new Date('2026-06-21T00:00:00.000Z');

const summary = workbook.addWorksheet('Summary');
summary.columns = [
  { width: 20 },
  { width: 28 },
  { width: 18 },
  { width: 18 },
  { width: 18 },
  { width: 18 },
];
summary.mergeCells('A1:F1');
summary.getCell('A1').value = 'Excel Template Engine - Production Showcase';
summary.getCell('A1').style = titleStyle();
summary.getCell('A3').value = 'Teacher';
summary.getCell('B3').value = '{{teacher.name}}';
summary.getCell('A4').value = 'Email';
summary.getCell('B4').value = '{{user.profile.email ?? "unknown@example.com"}}';
summary.getCell('A5').value = 'Contract';
summary.getCell('B5').value = '{{contract.code}}';
summary.getCell('A6').value = 'Total';
summary.getCell('B6').value = '{{currency(contract.total)}}';
summary.getCell('A7').value = 'Score Sum';
summary.getCell('B7').value = '{{sum(scores)}}';
summary.getCell('A9').value = 'Logo';
summary.getCell('B9').value = '{{image logo}}';
summary.getRow(9).height = 60;
summary.getCell('A11').value = 'Cross-sheet total formula';
summary.getCell('B11').value = { formula: 'SUM(Students!C3:C5)', result: 27 } as ExcelJS.CellValue;
styleLabelValue(summary, 3, 7);
summary.getCell('A11').style = labelStyle();
summary.getCell('B11').style = valueStyle();

const students = workbook.addWorksheet('Students');
students.columns = [
  { width: 8 },
  { width: 28 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
];
students.mergeCells('A1:E1');
students.getCell('A1').value = 'Students - Each Row + Formula';
students.getCell('A1').style = titleStyle();
students.getCell('A2').value = 'No.';
students.getCell('B2').value = 'Name';
students.getCell('C2').value = 'Score';
students.getCell('D2').value = 'Index';
students.getCell('E2').value = 'Flags';
students.getCell('F2').value = 'Score x 2';
for (const cellAddress of ['A2', 'B2', 'C2', 'D2', 'E2']) {
  students.getCell(cellAddress).style = headerStyle();
}
students.getCell('F2').style = headerStyle();
students.getCell('A3').value = '{{#block students}}';
students.getCell('A4').value = '{{code}}';
students.getCell('B4').value = '{{name}}';
students.getCell('C4').value = '{{score}}';
students.getCell('D4').value = '{{index}}';
students.getCell('E4').value = '{{first}} / {{last}}';
students.getCell('F4').value = { formula: 'C4*2', result: 16 } as ExcelJS.CellValue;
students.getCell('A5').value = '{{/block}}';
students.getRow(4).height = 24;
for (const cellAddress of ['A4', 'B4', 'C4', 'D4', 'E4', 'F4']) {
  students.getCell(cellAddress).style = valueStyle();
}

const loops = workbook.addWorksheet('Loops');
loops.columns = [
  { width: 36 },
  { width: 36 },
];
loops.mergeCells('A1:B1');
loops.getCell('A1').value = 'Loops - Each Metadata';
loops.getCell('A1').style = titleStyle();
loops.getCell('A2').value = 'Students';
loops.getCell('B2').value = 'Nested Scores';
loops.getCell('A2').style = headerStyle();
loops.getCell('B2').style = headerStyle();
loops.getCell('A3').value = 'Loop metadata: {{#each students}}{{index}}. {{name}} first={{first}} last={{last}}; {{/each}}';
loops.getCell('B3').value = 'Nested: {{#each classes}}{{name}}: {{#each students}}{{parent.index}}.{{index}} {{name}}; {{/each}}{{/each}}';
loops.getCell('A3').style = valueStyle();
loops.getCell('B3').style = valueStyle();

const matrix = workbook.addWorksheet('Matrix');
matrix.columns = [
  { width: 20 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
  { width: 14 },
];
matrix.mergeCells('A1:E1');
matrix.getCell('A1').value = 'Matrix - Each Column + Grid';
matrix.getCell('A1').style = titleStyle();
matrix.getCell('A2').value = 'Student / Subject';
matrix.getCell('B2').value = '{{#each-col subjects}}{{name}}{{/each-col}}';
matrix.getCell('A3').value = '{{#each students}}{{name}}{{/each}}';
matrix.getCell('B3').value = '{{#grid students subjects}}{{score}}{{/grid}}';
matrix.getCell('A2').style = headerStyle();
matrix.getCell('B2').style = headerStyle();
matrix.getCell('A3').style = valueStyle();
matrix.getCell('B3').style = valueStyle();

const dynamicHeader = workbook.addWorksheet('DynamicHeader');
dynamicHeader.columns = [
  { width: 20 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
];
dynamicHeader.mergeCells('A1:I1');
dynamicHeader.getCell('A1').value = 'Dynamic Grouped Header Test';
dynamicHeader.getCell('A1').style = titleStyle();
dynamicHeader.mergeCells('A2:A3');
dynamicHeader.getCell('A2').value = 'Category';
dynamicHeader.getCell('B2').value = '{{#each-col groups span=size reserve=8}}{{name}}{{/each-col}}';
dynamicHeader.getCell('B3').value = '{{#each-col groupCols reserve=8}}{{col}}{{/each-col}}';
dynamicHeader.getCell('A4').value = '{{#block rows}}';
dynamicHeader.getCell('A5').value = '{{label}}';
dynamicHeader.getCell('B5').value = '{{#each-col rowCols reserve=8}}{{value}}{{/each-col}}';
dynamicHeader.getCell('A6').value = '{{/block}}';
for (const cellAddress of ['A2', 'B2', 'B3', 'A5', 'B5']) {
  dynamicHeader.getCell(cellAddress).style = headerStyle();
}

const departmentReport = workbook.addWorksheet('DepartmentReport');
departmentReport.columns = [
  { width: 8 },
  { width: 26 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
  { width: 12 },
];
departmentReport.mergeCells('A1:J1');
departmentReport.getCell('A1').value = 'Department Overtime Report';
departmentReport.getCell('A1').style = titleStyle();
departmentReport.mergeCells('A2:A3');
departmentReport.getCell('A2').value = 'No.';
departmentReport.mergeCells('B2:B3');
departmentReport.getCell('B2').value = 'Lecturer';
departmentReport.getCell('C2').value = '{{#each-col semesterGroups span=colCount reserve=8}}{{name}}{{/each-col}}';
departmentReport.getCell('C3').value = '{{#each-col semesterColumns reserve=8}}{{name}}{{/each-col}}';
departmentReport.getCell('A4').value = '{{#block departments}}';
departmentReport.mergeCells('A5:J5');
departmentReport.getCell('A5').value = '{{index}}. {{name}}';
departmentReport.getCell('A6').value = '{{#block lecturers}}';
departmentReport.getCell('A7').value = '{{index}}';
departmentReport.getCell('B7').value = '{{name}}';
departmentReport.getCell('C7').value = '{{#each-col values reserve=8}}{{value}}{{/each-col}}';
departmentReport.getCell('A8').value = '{{/block}}';
departmentReport.getCell('A9').value = '{{/block}}';
for (const cellAddress of ['A2', 'B2', 'C2', 'C3', 'A5', 'A7', 'B7', 'C7']) {
  departmentReport.getCell(cellAddress).style = headerStyle();
}

const contracts = workbook.addWorksheet('Contracts');
contracts.columns = [
  { width: 20 },
  { width: 18 },
  { width: 18 },
  { width: 18 },
];
contracts.mergeCells('A1:D1');
contracts.getCell('A1').value = 'Contracts - Block Clone';
contracts.getCell('A1').style = titleStyle();
contracts.getCell('A2').value = '{{#block contracts}}';
contracts.mergeCells('A3:D3');
contracts.getCell('A3').value = 'Contract {{code}}';
contracts.getCell('A3').style = headerStyle();
contracts.getCell('A4').value = 'Title';
contracts.getCell('B4').value = '{{title}}';
contracts.getCell('C4').value = 'Amount';
contracts.getCell('D4').value = '{{amount}}';
contracts.getCell('A5').value = 'Amount x 2';
contracts.getCell('B5').value = { formula: 'D4*2', result: 2000 } as ExcelJS.CellValue;
contracts.getCell('C5').value = 'Teacher';
contracts.getCell('D5').value = '{{teacher.name ?? "Unknown"}}';
contracts.getCell('A6').value = '{{/block}}';
for (const rowNumber of [4, 5]) {
  contracts.getCell(`A${rowNumber}`).style = labelStyle();
  contracts.getCell(`C${rowNumber}`).style = labelStyle();
  contracts.getCell(`B${rowNumber}`).style = valueStyle();
  contracts.getCell(`D${rowNumber}`).style = valueStyle();
}

await workbook.xlsx.writeFile('examples/templates/production-showcase-template.xlsx');

function titleStyle(): Partial<ExcelJS.Style> {
  return {
    font: { name: 'Arial', bold: true, size: 16, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  };
}

function headerStyle(): Partial<ExcelJS.Style> {
  return {
    font: { name: 'Arial', bold: true, color: { argb: 'FF111827' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } },
    border: { bottom: { style: 'thin', color: { argb: 'FF94A3B8' } } },
    alignment: { horizontal: 'center', vertical: 'middle' },
  };
}

function labelStyle(): Partial<ExcelJS.Style> {
  return {
    font: { name: 'Arial', bold: true },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } },
  };
}

function valueStyle(): Partial<ExcelJS.Style> {
  return {
    font: { name: 'Arial' },
    alignment: { vertical: 'middle' },
  };
}

function styleLabelValue(worksheet: ExcelJS.Worksheet, startRow: number, endRow: number): void {
  for (let row = startRow; row <= endRow; row += 1) {
    worksheet.getCell(row, 1).style = labelStyle();
    worksheet.getCell(row, 2).style = valueStyle();
  }
}
