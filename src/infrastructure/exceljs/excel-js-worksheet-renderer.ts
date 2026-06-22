import type ExcelJS from 'exceljs';
import type { AssetResolver } from '../../application/managers/ports.js';
import type {
  CellDataValidation,
  CloneBlockOperation,
  CloneColumnOperation,
  CloneRowOperation,
  DeleteRowsOperation,
  ApplyMergeOperation,
  InsertImageOperation,
  RenderOperation,
  SetCellValueOperation,
} from '../../application/planner/render-plan.js';
import { ExcelJsBlockRenderer } from './excel-js-block-renderer.js';
import { ExcelJsFormulaManager } from './excel-js-formula-manager.js';
import { ExcelJsImageManager } from './excel-js-image-manager.js';
import { ExcelJsMergeManager } from './excel-js-merge-manager.js';
import { ExcelJsRowRenderer } from './excel-js-row-renderer.js';
import { ExcelJsStyleCloneManager } from './excel-js-style-clone-manager.js';
import { MergeRange } from '../../core/merge/merge-range.js';
import type { RangeAddress } from '../../shared/address/address.js';

export class ExcelJsWorksheetRenderer {
  constructor(
    private readonly worksheet: ExcelJS.Worksheet,
    private readonly assetResolver?: AssetResolver,
    private readonly options: ExcelJsWorksheetRendererOptions = {},
  ) {}

  async apply(operation: RenderOperation): Promise<void> {
    if (operation.type === 'CloneColumn') {
      await this.applyCloneColumn(operation);
      return;
    }

    if (operation.type === 'CloneRow') {
      await this.applyCloneRow(operation);
      return;
    }

    if (operation.type === 'CloneBlock') {
      await this.applyCloneBlock(operation);
      return;
    }

    if (operation.type === 'SetCellValue') {
      await this.applySetCellValue(operation);
      return;
    }

    if (operation.type === 'InsertImage') {
      await this.applyInsertImage(operation);
      return;
    }

    if (operation.type === 'DeleteRows') {
      await this.applyDeleteRows(operation);
      return;
    }

    if (operation.type === 'ApplyMerge') {
      await this.applyMerge(operation);
      return;
    }

    if (operation.type === 'CleanupTemplateMarkers') {
      await this.applyCleanupTemplateMarkers(operation.sheetName);
    }
  }

  async applyAll(operations: readonly RenderOperation[]): Promise<void> {
    for (const operation of operations) {
      if (operation.type === 'SetCellValue') {
        await this.applySetCellValue(operation);
        continue;
      }

      await this.apply(operation);
    }
  }

  private async applySetCellValue(operation: SetCellValueOperation): Promise<void> {
    if (operation.styleSource && !this.isSameCell(operation.styleSource, operation.cell)) {
      await new ExcelJsStyleCloneManager(this.worksheet.workbook).cloneCellStyle(
        this.withSheetName(operation.styleSource, operation.sheetName),
        this.withSheetName(operation.cell, operation.sheetName),
      );
    }

    const row = this.worksheet.getRow(operation.cell.row);
    new ExcelJsRowRenderer(row).applySetCellValue(operation);
    const cell = row.getCell(operation.cell.column);
    this.applyCellFormat(cell, operation);
    this.applyDataValidation(cell, operation.dataValidation);
  }

  private async applyCloneColumn(operation: CloneColumnOperation): Promise<void> {
    if (operation.count <= 0) {
      return;
    }

    const styleManager = new ExcelJsStyleCloneManager(this.worksheet.workbook);
    const mergeManager = new ExcelJsMergeManager(this.worksheet.workbook);
    const mergeRanges = await mergeManager.cloneForOperation(operation);
    if (operation.targetColumn <= this.worksheet.columnCount) {
      const insertedColumns = Array.from({ length: operation.count }, () => []);
      this.worksheet.spliceColumns(operation.targetColumn, 0, ...insertedColumns);
    }

    const shouldCloneColumnStyle = this.hasColumnStyle(operation.sourceColumn);
    const styledRows = this.collectStyledRows(operation.sourceColumn);
    for (let index = 0; index < operation.count; index += 1) {
      const targetColumn = operation.targetColumn + index;
      if (shouldCloneColumnStyle) {
        await styleManager.cloneColumnStyle(operation.sheetName, operation.sourceColumn, targetColumn);
      }

      for (const rowNumber of styledRows) {
        await styleManager.cloneCellStyle(
          { sheetName: operation.sheetName, row: rowNumber, column: operation.sourceColumn },
          { sheetName: operation.sheetName, row: rowNumber, column: targetColumn },
        );
      }
    }

    await mergeManager.validateAndApply(operation.sheetName, mergeRanges);
  }

  private async applyCloneRow(operation: CloneRowOperation): Promise<void> {
    if (operation.count <= 0) {
      return;
    }

    const styleManager = new ExcelJsStyleCloneManager(this.worksheet.workbook);
    const mergeManager = new ExcelJsMergeManager(this.worksheet.workbook);
    const mergeRanges = await mergeManager.cloneForOperation(operation);
    const shiftedRowStyles = this.snapshotShiftedRowStyles(operation.targetRow);
    if (operation.targetRow <= this.worksheet.rowCount) {
      const insertedRows = Array.from({ length: operation.count }, () => []);
      this.worksheet.spliceRows(operation.targetRow, 0, ...insertedRows);
    }
    this.restoreShiftedRowStyles(shiftedRowStyles, operation.count);

    const shouldCloneRowStyle = this.hasRowStyle(operation.sourceRow);
    const styledColumns = this.collectStyledColumns(operation.sourceRow);
    for (let index = 0; index < operation.count; index += 1) {
      const targetRow = operation.targetRow + index;
      if (shouldCloneRowStyle) {
        await styleManager.cloneRowStyle(operation.sheetName, operation.sourceRow, targetRow);
      }

      for (const columnNumber of styledColumns) {
        await styleManager.cloneCellStyle(
          { sheetName: operation.sheetName, row: operation.sourceRow, column: columnNumber },
          { sheetName: operation.sheetName, row: targetRow, column: columnNumber },
        );
      }
    }

    await mergeManager.validateAndApply(operation.sheetName, mergeRanges);
  }

  private async applyCloneBlock(operation: CloneBlockOperation): Promise<void> {
    await new ExcelJsBlockRenderer(this.worksheet).render(operation);
  }

  private async applyInsertImage(operation: InsertImageOperation): Promise<void> {
    await new ExcelJsImageManager(this.worksheet, this.assetResolver).insertImage(
      operation.source,
      operation.cell,
      {
        ...(operation.width ? { width: operation.width } : {}),
        ...(operation.height ? { height: operation.height } : {}),
      },
    );
  }

  private async applyDeleteRows(operation: DeleteRowsOperation): Promise<void> {
    if (operation.count <= 0) {
      return;
    }

    const startRow = Math.min(operation.startRow, Math.max(this.worksheet.rowCount, 1));
    const mergeManager = new ExcelJsMergeManager(this.worksheet.workbook);
    const mergeRanges = await mergeManager.collect(operation.sheetName);
    for (const range of mergeRanges) {
      this.worksheet.unMergeCells(MergeRange.fromRangeAddress(range).toA1());
    }

    const shiftedRowStyles = this.snapshotShiftedRowStyles(startRow + operation.count);
    this.worksheet.spliceRows(startRow, operation.count);
    this.restoreShiftedRowStyles(shiftedRowStyles, -operation.count);
    if (!this.options.preserveFormulas) {
      this.shiftFormulasAfterRowDelete(startRow, operation.count);
    }

    const shiftedMerges = mergeRanges
      .map((range) => this.shiftRangeAfterRowDelete(range, startRow, operation.count))
      .filter((range): range is RangeAddress => range !== undefined);
    await mergeManager.validateAndApply(operation.sheetName, shiftedMerges);
  }

  private async applyMerge(operation: ApplyMergeOperation): Promise<void> {
    const styleManager = new ExcelJsStyleCloneManager(this.worksheet.workbook);
    const source = operation.range.start;

    for (let row = operation.range.start.row; row <= operation.range.end.row; row += 1) {
      for (let column = operation.range.start.column; column <= operation.range.end.column; column += 1) {
        if (row === source.row && column === source.column) {
          continue;
        }

        await styleManager.cloneCellStyle(
          { ...(source.sheetName ? { sheetName: source.sheetName } : {}), row: source.row, column: source.column },
          { ...(source.sheetName ? { sheetName: source.sheetName } : {}), row, column },
        );
      }
    }

    await new ExcelJsMergeManager(this.worksheet.workbook).validateAndApply(operation.sheetName, [operation.range]);
  }

  private async applyCleanupTemplateMarkers(sheetName: string): Promise<void> {
    const rowsToDelete: number[] = [];
    const cellsToClear: ExcelJS.Cell[] = [];
    this.worksheet.eachRow((row, rowNumber) => {
      let hasMarker = false;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value !== 'string') {
          return;
        }

        if (/^\s*\{\{(?:#block\b|\/block\s*}})/.test(cell.value)) {
          hasMarker = true;
          return;
        }

        if (/\{\{[^}]+}}/.test(cell.value)) {
          cellsToClear.push(cell);
        }
      });

      if (hasMarker) {
        rowsToDelete.push(rowNumber);
      }
    });

    for (const cell of cellsToClear) {
      cell.value = null;
    }

    for (const rowNumber of rowsToDelete.sort((left, right) => right - left)) {
      await this.applyDeleteRows({
        id: `cleanup_template_marker_${sheetName}_${rowNumber}`,
        type: 'DeleteRows',
        sheetName,
        startRow: rowNumber,
        count: 1,
      });
    }
  }

  private shiftFormulasAfterRowDelete(startRow: number, count: number): void {
    const formulaManager = new ExcelJsFormulaManager();
    this.worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber < startRow) {
        return;
      }

      row.eachCell({ includeEmpty: true }, (cell) => {
        const value = cell.value;
        if (!value || typeof value !== 'object' || !('formula' in value) || typeof value.formula !== 'string') {
          return;
        }

        cell.value = {
          ...value,
          formula: formulaManager.shiftFormula(value.formula, -count, 0),
        } as ExcelJS.CellValue;
      });
    });
  }

  private shiftRangeAfterRowDelete(
    range: RangeAddress,
    startRow: number,
    count: number,
  ): RangeAddress | undefined {
    const deletedEndRow = startRow + count - 1;
    if (range.end.row < startRow) {
      return range;
    }

    if (range.start.row > deletedEndRow) {
      return MergeRange.fromRangeAddress(range).shift(-count, 0).toRangeAddress();
    }

    return undefined;
  }

  private hasRowStyle(rowNumber: number): boolean {
    const row = this.worksheet.getRow(rowNumber);
    return row.height !== undefined || this.hasStyleFields(row as Partial<ExcelJS.Style>);
  }

  private hasColumnStyle(columnNumber: number): boolean {
    const column = this.worksheet.getColumn(columnNumber);
    return column.width !== undefined || this.hasStyleFields(column as Partial<ExcelJS.Style>);
  }

  private collectStyledColumns(rowNumber: number): readonly number[] {
    const row = this.worksheet.getRow(rowNumber);
    const columns: number[] = [];
    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      if (this.hasStyle(cell.style)) {
        columns.push(columnNumber);
      }
    });
    return columns;
  }

  private collectStyledRows(columnNumber: number): readonly number[] {
    const rows: number[] = [];
    this.worksheet.getColumn(columnNumber).eachCell?.({ includeEmpty: true }, (cell, rowNumber) => {
      if (this.hasStyle(cell.style)) {
        rows.push(rowNumber);
      }
    });
    return rows;
  }

  private hasStyle(style: Partial<ExcelJS.Style> | undefined): boolean {
    if (!style) {
      return false;
    }

    return Object.values(style).some((value) => value !== undefined);
  }

  private hasStyleFields(style: Partial<ExcelJS.Style>): boolean {
    return STYLE_KEYS.some((key) => style[key] !== undefined);
  }

  private snapshotShiftedRowStyles(startRow: number): readonly RowStyleSnapshot[] {
    const snapshots: RowStyleSnapshot[] = [];
    for (let rowNumber = startRow; rowNumber <= this.worksheet.rowCount; rowNumber += 1) {
      const row = this.worksheet.getRow(rowNumber);
      const cells: CellStyleSnapshot[] = [];
      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        if (this.hasStyle(cell.style)) {
          cells.push({
            column: columnNumber,
            style: this.deepClone(cell.style),
          });
        }
      });

      snapshots.push({
        row: rowNumber,
        ...(row.height !== undefined ? { height: row.height } : {}),
        style: this.deepClone(this.pickStyleFields(row as Partial<ExcelJS.Style>)),
        cells,
      });
    }

    return snapshots;
  }

  private restoreShiftedRowStyles(snapshots: readonly RowStyleSnapshot[], rowDelta: number): void {
    for (const snapshot of snapshots) {
      const targetRowNumber = snapshot.row + rowDelta;
      if (targetRowNumber < 1) {
        continue;
      }

      const row = this.worksheet.getRow(targetRowNumber);
      this.applyStyleFields(row as Partial<ExcelJS.Style>, snapshot.style);
      if (snapshot.height !== undefined) {
        row.height = snapshot.height;
      }

      for (const cellSnapshot of snapshot.cells) {
        row.getCell(cellSnapshot.column).style = this.deepClone(cellSnapshot.style);
      }
    }
  }

  private pickStyleFields(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
    const picked: Partial<ExcelJS.Style> = {};
    for (const key of STYLE_KEYS) {
      if (style[key] !== undefined) {
        picked[key] = this.deepClone(style[key]) as never;
      }
    }

    return picked;
  }

  private applyStyleFields(target: Partial<ExcelJS.Style>, source: Partial<ExcelJS.Style>): void {
    for (const key of STYLE_KEYS) {
      if (source[key] === undefined) {
        delete target[key];
        continue;
      }

      target[key] = this.deepClone(source[key]) as never;
    }
  }

  private applyCellFormat(cell: ExcelJS.Cell, operation: SetCellValueOperation): void {
    if (operation.format?.numFmt !== undefined) {
      cell.numFmt = operation.format.numFmt;
    }

    if (operation.format?.wrapText !== undefined) {
      cell.alignment = {
        ...cell.alignment,
        wrapText: operation.format.wrapText,
      };
    }
  }

  private applyDataValidation(cell: ExcelJS.Cell, dataValidation: CellDataValidation | undefined): void {
    if (!dataValidation) {
      return;
    }

    if (dataValidation.type === 'choice') {
      cell.dataValidation = {
        type: 'list',
        allowBlank: dataValidation.allowBlank ?? true,
        formulae: [dataValidation.formula ?? this.toInlineChoiceFormula(dataValidation.values ?? [])],
        showErrorMessage: dataValidation.showErrorMessage ?? true,
        ...(dataValidation.errorTitle ? { errorTitle: dataValidation.errorTitle } : {}),
        ...(dataValidation.error ? { error: dataValidation.error } : {}),
        ...(dataValidation.showInputMessage !== undefined ? { showInputMessage: dataValidation.showInputMessage } : {}),
        ...(dataValidation.promptTitle ? { promptTitle: dataValidation.promptTitle } : {}),
        ...(dataValidation.prompt ? { prompt: dataValidation.prompt } : {}),
      };
    }
  }

  private toInlineChoiceFormula(values: readonly (string | number | boolean)[]): string {
    const list = values
      .map((value) => String(value).replace(/"/g, '""'))
      .join(',');
    return `"${list}"`;
  }

  private withSheetName(address: SetCellValueOperation['cell'], sheetName: string): SetCellValueOperation['cell'] {
    return {
      ...address,
      sheetName: address.sheetName ?? sheetName,
    };
  }

  private isSameCell(left: SetCellValueOperation['cell'], right: SetCellValueOperation['cell']): boolean {
    return left.sheetName === right.sheetName
      && left.row === right.row
      && left.column === right.column;
  }

  private deepClone<T>(value: T): T {
    if (value === undefined || value === null) {
      return value;
    }

    if (typeof structuredClone === 'function') {
      return structuredClone(value);
    }

    return JSON.parse(JSON.stringify(value)) as T;
  }
}

interface ExcelJsWorksheetRendererOptions {
  readonly preserveFormulas?: boolean;
}

interface RowStyleSnapshot {
  readonly row: number;
  readonly height?: number;
  readonly style: Partial<ExcelJS.Style>;
  readonly cells: readonly CellStyleSnapshot[];
}

interface CellStyleSnapshot {
  readonly column: number;
  readonly style: Partial<ExcelJS.Style>;
}

const STYLE_KEYS = ['numFmt', 'font', 'fill', 'border', 'alignment', 'protection'] as const;
