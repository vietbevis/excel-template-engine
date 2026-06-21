import type ExcelJS from 'exceljs';
import type { MergeManager } from '../../application/managers/ports.js';
import type { CloneBlockOperation, CloneColumnOperation, CloneRowOperation, RenderOperation } from '../../application/planner/render-plan.js';
import { MergeRange } from '../../core/merge/merge-range.js';
import { MergeTracker } from '../../core/merge/merge-tracker.js';
import type { RangeAddress } from '../../shared/address/address.js';

export class ExcelJsMergeManager implements MergeManager {
  constructor(private readonly workbook: ExcelJS.Workbook) {}

  async collect(sheetName: string): Promise<readonly RangeAddress[]> {
    return this.collectMergeRanges(sheetName).map((range) => range.toRangeAddress());
  }

  shift(range: RangeAddress, rowDelta: number, columnDelta: number): RangeAddress {
    return MergeRange.fromRangeAddress(range).shift(rowDelta, columnDelta).toRangeAddress();
  }

  async cloneForOperation(operation: RenderOperation): Promise<readonly RangeAddress[]> {
    if (operation.type === 'CloneRow') {
      return this.cloneForRow(operation).map((range) => range.toRangeAddress());
    }

    if (operation.type === 'CloneBlock') {
      return this.cloneForBlock(operation).map((range) => range.toRangeAddress());
    }

    if (operation.type === 'CloneColumn') {
      return this.cloneForColumn(operation).map((range) => range.toRangeAddress());
    }

    return [];
  }

  async validateAndApply(sheetName: string, ranges: readonly RangeAddress[]): Promise<void> {
    const existing = this.collectMergeRanges(sheetName);
    const tracker = new MergeTracker(existing);
    const worksheet = this.getWorksheet(sheetName);

    for (const range of ranges) {
      const mergeRange = MergeRange.fromRangeAddress(range);
      if (existing.some((current) => current.equals(mergeRange))) {
        continue;
      }
      tracker.add(mergeRange);
      worksheet.mergeCells(mergeRange.toA1());
    }
  }

  private cloneForRow(operation: CloneRowOperation): readonly MergeRange[] {
    const sourceMerges = this.collectMergeRanges(operation.sheetName)
      .filter((range) => range.start.row === operation.sourceRow && range.end.row === operation.sourceRow);
    const clones: MergeRange[] = [];

    for (let index = 0; index < operation.count; index += 1) {
      const targetRow = operation.targetRow + index;
      for (const range of sourceMerges) {
        clones.push(range.shift(targetRow - operation.sourceRow, 0));
      }
    }

    this.validateGenerated(operation.sheetName, clones);
    return clones;
  }

  private cloneForBlock(operation: CloneBlockOperation): readonly MergeRange[] {
    const sourceBlock = MergeRange.fromRangeAddress(operation.sourceRange);
    const sourceMerges = this.collectMergeRanges(operation.sheetName)
      .filter((range) => sourceBlock.containsRange(range));
    const blockHeight = sourceBlock.end.row - sourceBlock.start.row + 1;
    const blockWidth = sourceBlock.end.column - sourceBlock.start.column + 1;
    const clones: MergeRange[] = [];

    for (let index = 0; index < operation.count; index += 1) {
      const targetTopLeft = operation.direction === 'down'
        ? {
          ...operation.targetTopLeft,
          row: operation.targetTopLeft.row + index * blockHeight,
        }
        : {
          ...operation.targetTopLeft,
          column: operation.targetTopLeft.column + index * blockWidth,
        };

      for (const range of sourceMerges) {
        clones.push(range.cloneRelativeTo(sourceBlock, targetTopLeft));
      }
    }

    this.validateGenerated(operation.sheetName, clones);
    return clones;
  }

  private cloneForColumn(operation: CloneColumnOperation): readonly MergeRange[] {
    const sourceMerges = this.collectMergeRanges(operation.sheetName)
      .filter((range) => range.start.column === operation.sourceColumn && range.end.column === operation.sourceColumn);
    const clones: MergeRange[] = [];

    for (let index = 0; index < operation.count; index += 1) {
      const targetColumn = operation.targetColumn + index;
      for (const range of sourceMerges) {
        clones.push(range.shift(0, targetColumn - operation.sourceColumn));
      }
    }

    this.validateGenerated(operation.sheetName, clones);
    return clones;
  }

  private validateGenerated(sheetName: string, ranges: readonly MergeRange[]): void {
    const tracker = new MergeTracker(this.collectMergeRanges(sheetName));
    ranges.forEach((range) => tracker.add(range));
  }

  private collectMergeRanges(sheetName: string): readonly MergeRange[] {
    const worksheet = this.getWorksheet(sheetName);
    const rawMerges = (worksheet as ExcelJsWorksheetWithMerges)._merges ?? {};
    const unique = new Map<string, MergeRange>();

    Object.values(rawMerges).forEach((merge) => {
      const range = MergeRange.fromA1(merge.range, sheetName);
      unique.set(range.toA1(), range);
    });

    return [...unique.values()];
  }

  private getWorksheet(sheetName: string): ExcelJS.Worksheet {
    const worksheet = this.workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Worksheet not found: ${sheetName}`);
    }

    return worksheet;
  }
}

interface ExcelJsWorksheetWithMerges extends ExcelJS.Worksheet {
  readonly _merges?: Record<string, { readonly range: string }>;
}
