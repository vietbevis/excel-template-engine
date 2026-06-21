import type { CellAddress } from '../../shared/address/address.js';

export class RenderCursor {
  constructor(
    readonly origin: CellAddress,
  ) {}

  at(rowOffset: number, columnOffset: number): CellAddress {
    return {
      ...(this.origin.sheetName ? { sheetName: this.origin.sheetName } : {}),
      row: this.origin.row + rowOffset,
      column: this.origin.column + columnOffset,
    };
  }
}
