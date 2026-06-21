import type ExcelJS from 'exceljs';
import type { CloneBlockOperation } from '../../application/planner/render-plan.js';
import type { CellAddress, RangeAddress } from '../../shared/address/address.js';
import { ExcelJsFormulaManager } from './excel-js-formula-manager.js';
import { ExcelJsMergeManager } from './excel-js-merge-manager.js';
import { ExcelJsStyleCloneManager } from './excel-js-style-clone-manager.js';

export class ExcelJsBlockCloneManager {
  constructor(private readonly worksheet: ExcelJS.Worksheet) {}

  async clone(operation: CloneBlockOperation): Promise<void> {
    if (operation.count <= 0) {
      return;
    }

    if (operation.direction !== 'down') {
      throw new Error('CloneBlock direction "right" is not supported by ExcelJS block renderer yet.');
    }

    const blockHeight = this.getBlockHeight(operation.sourceRange);
    const rowsToInsert = blockHeight * operation.count;
    const snapshots = this.createSnapshots(operation);
    const images = this.collectSourceImages(operation.sourceRange);
    const mergeManager = new ExcelJsMergeManager(this.worksheet.workbook);
    const mergeRanges = await mergeManager.cloneForOperation(operation);

    this.worksheet.spliceRows(operation.targetTopLeft.row, 0, ...Array.from({ length: rowsToInsert }, () => []));

    const styleManager = new ExcelJsStyleCloneManager(this.worksheet.workbook);
    await this.cloneRows(operation, styleManager);
    await this.cloneCells(operation, snapshots, styleManager);
    this.cloneImages(operation, images);
    await mergeManager.validateAndApply(operation.sheetName, mergeRanges);
  }

  private createSnapshots(operation: CloneBlockOperation): readonly CellSnapshot[] {
    const snapshots: CellSnapshot[] = [];

    for (let row = operation.sourceRange.start.row; row <= operation.sourceRange.end.row; row += 1) {
      for (let column = operation.sourceRange.start.column; column <= operation.sourceRange.end.column; column += 1) {
        const cell = this.worksheet.getCell(row, column);
        snapshots.push({
          rowOffset: row - operation.sourceRange.start.row,
          columnOffset: column - operation.sourceRange.start.column,
          value: this.deepClone(cell.value),
        });
      }
    }

    return snapshots;
  }

  private async cloneRows(
    operation: CloneBlockOperation,
    styleManager: ExcelJsStyleCloneManager,
  ): Promise<void> {
    const blockHeight = this.getBlockHeight(operation.sourceRange);

    for (let cloneIndex = 0; cloneIndex < operation.count; cloneIndex += 1) {
      for (let rowOffset = 0; rowOffset < blockHeight; rowOffset += 1) {
        await styleManager.cloneRowStyle(
          operation.sheetName,
          operation.sourceRange.start.row + rowOffset,
          operation.targetTopLeft.row + cloneIndex * blockHeight + rowOffset,
        );
      }
    }
  }

  private async cloneCells(
    operation: CloneBlockOperation,
    snapshots: readonly CellSnapshot[],
    styleManager: ExcelJsStyleCloneManager,
  ): Promise<void> {
    const blockHeight = this.getBlockHeight(operation.sourceRange);
    const formulaManager = new ExcelJsFormulaManager();

    for (let cloneIndex = 0; cloneIndex < operation.count; cloneIndex += 1) {
      for (const snapshot of snapshots) {
        const source: CellAddress = {
          sheetName: operation.sheetName,
          row: operation.sourceRange.start.row + snapshot.rowOffset,
          column: operation.sourceRange.start.column + snapshot.columnOffset,
        };
        const target: CellAddress = {
          sheetName: operation.sheetName,
          row: operation.targetTopLeft.row + cloneIndex * blockHeight + snapshot.rowOffset,
          column: operation.targetTopLeft.column + snapshot.columnOffset,
        };
        const rowDelta = target.row - source.row;
        const columnDelta = target.column - source.column;

        this.worksheet.getCell(target.row, target.column).value = this.shiftFormulaValue(
          snapshot.value,
          rowDelta,
          columnDelta,
          formulaManager,
        );
        await styleManager.cloneCellStyle(source, target);
      }
    }
  }

  private collectSourceImages(sourceRange: RangeAddress): readonly ImageSnapshot[] {
    return this.worksheet.getImages()
      .filter((image) => this.imageIsInsideRange(image, sourceRange))
      .map((image) => ({
        imageId: Number(image.imageId),
        range: this.deepClone(image.range),
      }));
  }

  private cloneImages(operation: CloneBlockOperation, images: readonly ImageSnapshot[]): void {
    const blockHeight = this.getBlockHeight(operation.sourceRange);

    for (let cloneIndex = 0; cloneIndex < operation.count; cloneIndex += 1) {
      const rowDelta = operation.targetTopLeft.row - operation.sourceRange.start.row + cloneIndex * blockHeight;
      const columnDelta = operation.targetTopLeft.column - operation.sourceRange.start.column;

      for (const image of images) {
        this.worksheet.addImage(image.imageId, this.shiftImageRange(image.range, rowDelta, columnDelta));
      }
    }
  }

  private imageIsInsideRange(image: WorksheetImage, sourceRange: RangeAddress): boolean {
    const range = image.range as ImageRangeLike;
    if (!range.tl || !range.br) {
      return false;
    }

    const startColumn = sourceRange.start.column - 1;
    const endColumn = sourceRange.end.column;
    const startRow = sourceRange.start.row - 1;
    const endRow = sourceRange.end.row;

    return range.tl.col >= startColumn
      && range.br.col <= endColumn
      && range.tl.row >= startRow
      && range.br.row <= endRow;
  }

  private shiftImageRange(range: ExcelJS.ImageRange, rowDelta: number, columnDelta: number): ExcelJS.ImageRange {
    const clone = this.deepClone(range) as ImageRangeLike;

    if (clone.tl) {
      clone.tl.col += columnDelta;
      clone.tl.row += rowDelta;
    }

    if (clone.br) {
      clone.br.col += columnDelta;
      clone.br.row += rowDelta;
    }

    return clone as ExcelJS.ImageRange;
  }

  private getBlockHeight(sourceRange: RangeAddress): number {
    return sourceRange.end.row - sourceRange.start.row + 1;
  }

  private shiftFormulaValue(
    value: ExcelJS.CellValue,
    rowDelta: number,
    columnDelta: number,
    formulaManager: ExcelJsFormulaManager,
  ): ExcelJS.CellValue {
    const clone = this.deepClone(value);
    if (!clone || typeof clone !== 'object' || !('formula' in clone) || typeof clone.formula !== 'string') {
      return clone;
    }

    const { result: _result, ...formulaValue } = clone;
    return {
      ...formulaValue,
      formula: formulaManager.shiftFormula(clone.formula, rowDelta, columnDelta),
    } as ExcelJS.CellValue;
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

interface CellSnapshot {
  readonly rowOffset: number;
  readonly columnOffset: number;
  readonly value: ExcelJS.CellValue;
}

interface ImageSnapshot {
  readonly imageId: number;
  readonly range: ExcelJS.ImageRange;
}

interface ImageRangeLike {
  tl?: { col: number; row: number };
  br?: { col: number; row: number };
}

interface WorksheetImage {
  readonly imageId: string;
  readonly range: ExcelJS.ImageRange;
}
