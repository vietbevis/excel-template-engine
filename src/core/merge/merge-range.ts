import type { CellAddress, RangeAddress } from '../../shared/address/address.js';
import { AddressParser } from '../../shared/address/address-parser.js';

const addressParser = new AddressParser();

export class MergeRange {
  readonly sheetName: string | undefined;
  readonly start: CellAddress;
  readonly end: CellAddress;

  private constructor(range: RangeAddress) {
    const normalized = MergeRange.normalize(range);
    this.sheetName = normalized.sheetName;
    this.start = normalized.start;
    this.end = normalized.end;
  }

  static fromRangeAddress(range: RangeAddress): MergeRange {
    return new MergeRange(range);
  }

  static fromA1(ref: string, sheetName?: string): MergeRange {
    const parsed = addressParser.parseRange(ref);
    const finalSheetName = parsed.sheetName ?? sheetName;

    return new MergeRange(finalSheetName ? {
      sheetName: finalSheetName,
      start: { ...parsed.start, sheetName: finalSheetName },
      end: { ...parsed.end, sheetName: finalSheetName },
    } : parsed);
  }

  toRangeAddress(): RangeAddress {
    return this.sheetName ? {
      sheetName: this.sheetName,
      start: { ...this.start, sheetName: this.sheetName },
      end: { ...this.end, sheetName: this.sheetName },
    } : {
      start: this.start,
      end: this.end,
    };
  }

  toA1(): string {
    return `${this.cellToA1(this.start)}:${this.cellToA1(this.end)}`;
  }

  intersects(other: MergeRange): boolean {
    if (this.sheetName && other.sheetName && this.sheetName !== other.sheetName) {
      return false;
    }

    return !(
      this.end.row < other.start.row
      || other.end.row < this.start.row
      || this.end.column < other.start.column
      || other.end.column < this.start.column
    );
  }

  equals(other: MergeRange): boolean {
    return (this.sheetName ?? '') === (other.sheetName ?? '')
      && this.start.row === other.start.row
      && this.start.column === other.start.column
      && this.end.row === other.end.row
      && this.end.column === other.end.column;
  }

  containsRange(other: MergeRange): boolean {
    if (this.sheetName && other.sheetName && this.sheetName !== other.sheetName) {
      return false;
    }

    return this.start.row <= other.start.row
      && this.start.column <= other.start.column
      && this.end.row >= other.end.row
      && this.end.column >= other.end.column;
  }

  shift(rowDelta: number, columnDelta: number): MergeRange {
    return new MergeRange({
      ...(this.sheetName ? { sheetName: this.sheetName } : {}),
      start: {
        ...(this.sheetName ? { sheetName: this.sheetName } : {}),
        row: this.start.row + rowDelta,
        column: this.start.column + columnDelta,
      },
      end: {
        ...(this.sheetName ? { sheetName: this.sheetName } : {}),
        row: this.end.row + rowDelta,
        column: this.end.column + columnDelta,
      },
    });
  }

  cloneRelativeTo(sourceRange: MergeRange, targetTopLeft: CellAddress): MergeRange {
    const rowDelta = targetTopLeft.row - sourceRange.start.row;
    const columnDelta = targetTopLeft.column - sourceRange.start.column;
    const sheetName = targetTopLeft.sheetName ?? this.sheetName;
    const shifted = this.shift(rowDelta, columnDelta);

    return sheetName ? new MergeRange({
      sheetName,
      start: { ...shifted.start, sheetName },
      end: { ...shifted.end, sheetName },
    }) : shifted;
  }

  private static normalize(range: RangeAddress): RangeAddress {
    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startColumn = Math.min(range.start.column, range.end.column);
    const endColumn = Math.max(range.start.column, range.end.column);
    const sheetName = range.sheetName ?? range.start.sheetName ?? range.end.sheetName;

    return sheetName ? {
      sheetName,
      start: { sheetName, row: startRow, column: startColumn },
      end: { sheetName, row: endRow, column: endColumn },
    } : {
      start: { row: startRow, column: startColumn },
      end: { row: endRow, column: endColumn },
    };
  }

  private cellToA1(cell: CellAddress): string {
    return `${addressParser.columnNumberToName(cell.column)}${cell.row}`;
  }
}
