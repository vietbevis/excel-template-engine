import type ExcelJS from 'exceljs';
import type { AssetResolver } from '../../application/managers/ports.js';
import type {
  CloneBlockOperation,
  CloneColumnOperation,
  CloneRowOperation,
  InsertImageOperation,
  RenderOperation,
  SetCellValueOperation,
} from '../../application/planner/render-plan.js';
import { ExcelJsBlockRenderer } from './excel-js-block-renderer.js';
import { ExcelJsImageManager } from './excel-js-image-manager.js';
import { ExcelJsMergeManager } from './excel-js-merge-manager.js';
import { ExcelJsRowRenderer } from './excel-js-row-renderer.js';
import { ExcelJsStyleCloneManager } from './excel-js-style-clone-manager.js';

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
    const insertedColumns = Array.from({ length: operation.count }, () => []);
    this.worksheet.spliceColumns(operation.targetColumn, 0, ...insertedColumns);

    for (let index = 0; index < operation.count; index += 1) {
      const targetColumn = operation.targetColumn + index;
      await styleManager.cloneColumnStyle(operation.sheetName, operation.sourceColumn, targetColumn);

      for (let rowNumber = 1; rowNumber <= this.worksheet.rowCount; rowNumber += 1) {
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
    const insertedRows = Array.from({ length: operation.count }, () => []);
    this.worksheet.spliceRows(operation.targetRow, 0, ...insertedRows);

    for (let index = 0; index < operation.count; index += 1) {
      const targetRow = operation.targetRow + index;
      await styleManager.cloneRowStyle(operation.sheetName, operation.sourceRow, targetRow);

      for (let columnNumber = 1; columnNumber <= this.worksheet.columnCount; columnNumber += 1) {
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
}
