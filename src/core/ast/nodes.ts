import type { CellAddress, RangeAddress } from '../../shared/address/address.js';

export interface WorkbookAST {
  readonly kind: 'WorkbookAST';
  readonly sheets: readonly SheetAST[];
}

export interface SheetAST {
  readonly kind: 'SheetAST';
  readonly name: string;
  readonly sheetId: string;
  readonly rows: readonly RowAST[];
}

export interface RowAST {
  readonly kind: 'RowAST';
  readonly rowNumber: number;
  readonly cells: readonly CellAST[];
}

export interface CellAST {
  readonly kind: 'CellAST';
  readonly address: CellAddress;
  readonly sourceRange?: RangeAddress;
  readonly nodes: readonly TemplateNode[];
  readonly rawValue: unknown;
}

export type TemplateNode =
  | TextNode
  | PlaceholderNode
  | EachNode
  | EachColumnNode
  | GridNode
  | BlockNode
  | BlockStartNode
  | BlockEndNode
  | IfNode
  | HelperNode
  | ImageNode;

export interface BaseNode {
  readonly id: string;
  readonly source: SourceSpan;
}

export interface SourceSpan {
  readonly start: number;
  readonly end: number;
}

export interface TextNode extends BaseNode {
  readonly kind: 'TextNode';
  readonly value: string;
}

export interface PlaceholderNode extends BaseNode {
  readonly kind: 'PlaceholderNode';
  readonly path: string;
}

export interface EachNode extends BaseNode {
  readonly kind: 'EachNode';
  readonly path: string;
  readonly children: readonly TemplateNode[];
}

export interface EachColumnNode extends BaseNode {
  readonly kind: 'EachColumnNode';
  readonly path: string;
  readonly spanPath?: string;
  readonly renderPath?: string;
  readonly rowSpanPath?: string;
  readonly rowSpan?: number;
  readonly reservedColumnsPath?: string;
  readonly reservedColumns?: number;
  readonly children: readonly TemplateNode[];
}

export interface GridNode extends BaseNode {
  readonly kind: 'GridNode';
  readonly rowPath: string;
  readonly columnPath: string;
  readonly children: readonly TemplateNode[];
}

export interface BlockNode extends BaseNode {
  readonly kind: 'BlockNode';
  readonly path: string;
  readonly children: readonly TemplateNode[];
}

export interface BlockStartNode extends BaseNode {
  readonly kind: 'BlockStartNode';
  readonly path: string;
}

export interface BlockEndNode extends BaseNode {
  readonly kind: 'BlockEndNode';
}

export interface IfNode extends BaseNode {
  readonly kind: 'IfNode';
  readonly conditionPath: string;
  readonly children: readonly TemplateNode[];
}

export interface HelperNode extends BaseNode {
  readonly kind: 'HelperNode';
  readonly name: string;
  readonly args: readonly HelperArgument[];
}

export interface ImageNode extends BaseNode {
  readonly kind: 'ImageNode';
  readonly path: string;
  readonly options: ImageNodeOptions;
}

export type HelperArgument =
  | { readonly kind: 'path'; readonly value: string }
  | { readonly kind: 'literal'; readonly value: string | number | boolean | null };

export interface ImageNodeOptions {
  readonly width?: number;
  readonly height?: number;
  readonly fit?: 'cell' | 'merge' | 'explicit';
}
