export interface CellAddress {
  readonly sheetName?: string;
  readonly row: number;
  readonly column: number;
}

export interface RangeAddress {
  readonly sheetName?: string;
  readonly start: CellAddress;
  readonly end: CellAddress;
}

export interface A1Reference {
  readonly sheetName?: string;
  readonly columnName: string;
  readonly column: number;
  readonly row: number;
  readonly absoluteColumn: boolean;
  readonly absoluteRow: boolean;
}
