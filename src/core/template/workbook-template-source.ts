import type { CellAddress, RangeAddress } from '../../shared/address/address.js';

export interface WorkbookTemplateSource {
  readonly sheets: readonly WorksheetTemplateSource[];
}

export interface WorksheetTemplateSource {
  readonly name: string;
  readonly sheetId: string;
  readonly rows: readonly RowTemplateSource[];
  readonly merges: readonly RangeAddress[];
}

export interface RowTemplateSource {
  readonly rowNumber: number;
  readonly height?: number;
  readonly cells: readonly CellTemplateSource[];
}

export interface CellTemplateSource {
  readonly address: CellAddress;
  readonly value: unknown;
  readonly formula?: string;
  readonly styleRef?: string;
}
