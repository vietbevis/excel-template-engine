import type { CellAddress, RangeAddress } from '../../shared/address/address.js';

export interface RenderPlan {
  readonly operations: readonly RenderOperation[];
  readonly warnings: readonly string[];
}

export type RenderOperation =
  | SetCellValueOperation
  | CloneRowOperation
  | CloneColumnOperation
  | CloneBlockOperation
  | InsertImageOperation
  | DeleteRowsOperation
  | CleanupTemplateMarkersOperation
  | ApplyMergeOperation
  | ShiftFormulaOperation
  | ClearFormulaCacheOperation;

export interface BaseOperation {
  readonly id: string;
  readonly sheetName: string;
}

export interface SetCellValueOperation extends BaseOperation {
  readonly type: 'SetCellValue';
  readonly cell: CellAddress;
  readonly value: unknown;
  readonly styleSource?: CellAddress;
  readonly format?: CellFormat;
  readonly dataValidation?: CellDataValidation;
}

export interface CellFormat {
  readonly wrapText?: boolean;
  readonly numFmt?: string;
}

export interface CellDataValidation {
  readonly type: 'choice';
  readonly values?: readonly (string | number | boolean)[];
  readonly formula?: string;
  readonly allowBlank?: boolean;
  readonly showErrorMessage?: boolean;
  readonly errorTitle?: string;
  readonly error?: string;
  readonly showInputMessage?: boolean;
  readonly promptTitle?: string;
  readonly prompt?: string;
}

export interface CloneRowOperation extends BaseOperation {
  readonly type: 'CloneRow';
  readonly sourceRow: number;
  readonly targetRow: number;
  readonly count: number;
}

export interface CloneColumnOperation extends BaseOperation {
  readonly type: 'CloneColumn';
  readonly sourceColumn: number;
  readonly targetColumn: number;
  readonly count: number;
}

export interface CloneBlockOperation extends BaseOperation {
  readonly type: 'CloneBlock';
  readonly sourceRange: RangeAddress;
  readonly targetTopLeft: CellAddress;
  readonly count: number;
  readonly direction: 'down' | 'right';
}

export interface InsertImageOperation extends BaseOperation {
  readonly type: 'InsertImage';
  readonly cell: CellAddress;
  readonly source: unknown;
  readonly width?: number;
  readonly height?: number;
}

export interface DeleteRowsOperation extends BaseOperation {
  readonly type: 'DeleteRows';
  readonly startRow: number;
  readonly count: number;
}

export interface CleanupTemplateMarkersOperation extends BaseOperation {
  readonly type: 'CleanupTemplateMarkers';
}

export interface ApplyMergeOperation extends BaseOperation {
  readonly type: 'ApplyMerge';
  readonly range: RangeAddress;
}

export interface ShiftFormulaOperation extends BaseOperation {
  readonly type: 'ShiftFormula';
  readonly range: RangeAddress;
  readonly rowDelta: number;
  readonly columnDelta: number;
}

export interface ClearFormulaCacheOperation extends BaseOperation {
  readonly type: 'ClearFormulaCache';
  readonly range?: RangeAddress;
}
