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
    const shiftedRowStyles = this.snapshotShiftedRowStyles(operation.targetTopLeft.row);

    this.worksheet.spliceRows(operation.targetTopLeft.row, 0, ...Array.from({ length: rowsToInsert }, () => []));
    this.restoreShiftedRowStyles(shiftedRowStyles, rowsToInsert);

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

  private hasStyle(style: Partial<ExcelJS.Style> | undefined): boolean {
    return !!style && STYLE_KEYS.some((key) => style[key] !== undefined);
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

const STYLE_KEYS = ['numFmt', 'font', 'fill', 'border', 'alignment', 'protection'] as const;
