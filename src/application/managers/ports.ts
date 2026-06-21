import type { RenderPlan, RenderOperation } from '../planner/render-plan.js';
import type { CellAddress, RangeAddress } from '../../shared/address/address.js';
import type { WorkbookTemplateSource } from '../../core/template/workbook-template-source.js';

export type { CellTemplateSource, RowTemplateSource, WorkbookTemplateSource, WorksheetTemplateSource } from '../../core/template/workbook-template-source.js';

export type TemplateInput = Buffer | Uint8Array | ArrayBuffer | string;

export interface WorkbookRenderer {
  load(input: TemplateInput): Promise<WorkbookTemplateSource>;
  apply(plan: RenderPlan): Promise<void>;
  write(): Promise<Uint8Array>;
}

export interface StyleCloneManager {
  cloneCellStyle(source: CellAddress, target: CellAddress): Promise<void>;
  cloneRowStyle(sheetName: string, sourceRow: number, targetRow: number): Promise<void>;
  cloneColumnStyle(sheetName: string, sourceColumn: number, targetColumn: number): Promise<void>;
}

export interface MergeManager {
  collect(sheetName: string): Promise<readonly RangeAddress[]>;
  shift(range: RangeAddress, rowDelta: number, columnDelta: number): RangeAddress;
  cloneForOperation(operation: RenderOperation): Promise<readonly RangeAddress[]>;
}

export interface FormulaManager {
  shiftFormula(formula: string, rowDelta: number, columnDelta: number): string;
  clearCachedValues(range?: RangeAddress): Promise<void>;
}

export interface ImageManager {
  insertImage(source: unknown, target: CellAddress, options?: ImageInsertOptions): Promise<void>;
}

export interface ImageInsertOptions {
  readonly width?: number;
  readonly height?: number;
  readonly fit?: 'cell' | 'merge' | 'explicit';
}

export interface AssetResolver {
  resolve(source: unknown): Promise<ResolvedAsset>;
}

export interface AssetResolverOptions {
  readonly baseDir?: string;
  readonly allowAbsolutePaths?: boolean;
  readonly maxBytes?: number;
}

export interface ResolvedAsset {
  readonly bytes: Uint8Array;
  readonly extension: 'png' | 'jpg' | 'jpeg' | 'gif';
  readonly contentType: string;
}
