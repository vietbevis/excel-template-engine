import type ExcelJS from 'exceljs';
import type { AssetResolver } from '../../application/managers/ports.js';
import type {
  CloneBlockOperation,
  CloneColumnOperation,
  CloneRowOperation,
  DeleteRowsOperation,
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
      this.applySetCellValue(operation);
      return;
    }

    if (operation.type === 'InsertImage') {
      await this.applyInsertImage(operation);
      return;
    }

    if (operation.type === 'DeleteRows') {
      await this.applyDeleteRows(operation);
    }
  }

  async applyAll(operations: readonly RenderOperation[]): Promise<void> {
    for (const operation of operations) {
      if (operation.type === 'SetCellValue') {
        this.applySetCellValue(operation);
        continue;
      }

      await this.apply(operation);
    }
  }

  private applySetCellValue(operation: SetCellValueOperation): void {
    const row = this.worksheet.getRow(operation.cell.row);
    new ExcelJsRowRenderer(row).applySetCellValue(operation);
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
    if (operation.targetRow <= this.worksheet.rowCount) {
      const insertedRows = Array.from({ length: operation.count }, () => []);
      this.worksheet.spliceRows(operation.targetRow, 0, ...insertedRows);
    }

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

    const mergeManager = new ExcelJsMergeManager(this.worksheet.workbook);
    const mergeRanges = await mergeManager.collect(operation.sheetName);
    for (const range of mergeRanges) {
      this.worksheet.unMergeCells(MergeRange.fromRangeAddress(range).toA1());
    }

    this.worksheet.spliceRows(operation.startRow, operation.count);
    this.shiftFormulasAfterRowDelete(operation.startRow, operation.count);

    const shiftedMerges = mergeRanges
      .map((range) => this.shiftRangeAfterRowDelete(range, operation.startRow, operation.count))
      .filter((range): range is RangeAddress => range !== undefined);
    await mergeManager.validateAndApply(operation.sheetName, shiftedMerges);
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
}

const STYLE_KEYS = ['numFmt', 'font', 'fill', 'border', 'alignment', 'protection'] as const;
