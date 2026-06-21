import type { A1Reference, RangeAddress } from './address.js';

export class AddressParser {
  parseCell(ref: string): A1Reference {
    const trimmed = ref.trim();
    const match = /^(?:(?<sheet>(?:'[^']+'|[^!]+))!)?(?<absCol>\$?)(?<col>[A-Za-z]+)(?<absRow>\$?)(?<row>\d+)$/.exec(trimmed);

    if (!match?.groups) {
      throw new Error(`Invalid Excel cell reference: ${ref}`);
    }

    const rawColumnName = match.groups.col;
    const rawRow = match.groups.row;

    if (!rawColumnName || !rawRow) {
      throw new Error(`Invalid Excel cell reference: ${ref}`);
    }

    const columnName = rawColumnName.toUpperCase();
    const sheetName = this.normalizeSheetName(match.groups.sheet);

    return this.withOptionalSheetName({
      columnName,
      column: this.columnNameToNumber(columnName),
      row: Number(rawRow),
      absoluteColumn: match.groups.absCol === '$',
      absoluteRow: match.groups.absRow === '$',
    }, sheetName);
  }

  parseRange(ref: string): RangeAddress {
    const [startRef, endRef] = ref.split(':');

    if (!startRef || !endRef) {
      throw new Error(`Invalid Excel range reference: ${ref}`);
    }

    const start = this.parseCell(startRef);
    const end = this.parseCell(endRef.includes('!') ? endRef : `${start.sheetName ? `${this.quoteSheetName(start.sheetName)}!` : ''}${endRef}`);
    const sheetName = start.sheetName ?? end.sheetName;

    return this.withOptionalRangeSheetName({
      start: this.withOptionalCellSheetName({
        row: start.row,
        column: start.column,
      }, sheetName),
      end: this.withOptionalCellSheetName({
        row: end.row,
        column: end.column,
      }, sheetName),
    }, sheetName);
  }

  columnNameToNumber(columnName: string): number {
    const normalized = columnName.trim().toUpperCase();

    if (!/^[A-Z]+$/.test(normalized)) {
      throw new Error(`Invalid Excel column name: ${columnName}`);
    }

    return normalized.split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
  }

  columnNumberToName(column: number): string {
    if (!Number.isInteger(column) || column < 1) {
      throw new Error(`Invalid Excel column number: ${column}`);
    }

    let value = column;
    let name = '';

    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }

    return name;
  }

  private normalizeSheetName(sheetName: string | undefined): string | undefined {
    if (!sheetName) {
      return undefined;
    }

    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
      return sheetName.slice(1, -1).replace(/''/g, "'");
    }

    return sheetName;
  }

  private quoteSheetName(sheetName: string): string {
    return /^[A-Za-z0-9_]+$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
  }

  private withOptionalSheetName(reference: Omit<A1Reference, 'sheetName'>, sheetName: string | undefined): A1Reference {
    return sheetName ? { ...reference, sheetName } : reference;
  }

  private withOptionalCellSheetName<T extends { readonly row: number; readonly column: number }>(
    cell: T,
    sheetName: string | undefined,
  ): T & { readonly sheetName?: string } {
    return sheetName ? { ...cell, sheetName } : cell;
  }

  private withOptionalRangeSheetName(
    range: Omit<RangeAddress, 'sheetName'>,
    sheetName: string | undefined,
  ): RangeAddress {
    return sheetName ? { ...range, sheetName } : range;
  }
}
