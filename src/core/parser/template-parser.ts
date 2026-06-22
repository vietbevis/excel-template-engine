import type {
  BlockEndNode,
  BlockNode,
  BlockStartNode,
  EachColumnNode,
  EachNode,
  GridNode,
  HelperArgument,
  HelperNode,
  IfNode,
  ImageNode,
  PlaceholderNode,
  TemplateNode,
  TextNode,
  WorkbookAST,
} from '../ast/nodes.js';
import type { WorkbookTemplateSource } from '../template/workbook-template-source.js';

export class TemplateParser {
  parseWorkbook(source: WorkbookTemplateSource): WorkbookAST {
    return {
      kind: 'WorkbookAST',
      sheets: source.sheets.map((sheet) => ({
        kind: 'SheetAST',
        name: sheet.name,
        sheetId: sheet.sheetId,
        rows: sheet.rows.map((row) => ({
          kind: 'RowAST',
          rowNumber: row.rowNumber,
          cells: row.cells.map((cell) => ({
            kind: 'CellAST',
            address: cell.address,
            ...(cell.sourceRange ? { sourceRange: cell.sourceRange } : {}),
            nodes: typeof cell.value === 'string' ? this.parseCell(cell.value) : [],
            rawValue: cell.value,
          })),
        })),
      })),
    };
  }

  parseCell(input: string): readonly TemplateNode[] {
    const blockStart = /^\s*\{\{#block\s+([^{}]+)}}\s*$/.exec(input);
    if (blockStart?.[1]) {
      const node: BlockStartNode = {
        kind: 'BlockStartNode',
        id: this.createId('block-start', 0),
        path: blockStart[1].trim(),
        source: { start: 0, end: input.length },
      };
      return [node];
    }

    if (/^\s*\{\{\/block}}\s*$/.test(input)) {
      const node: BlockEndNode = {
        kind: 'BlockEndNode',
        id: this.createId('block-end', 0),
        source: { start: 0, end: input.length },
      };
      return [node];
    }

    return this.parseTemplate(input, 0, undefined).nodes;
  }

  private parseTemplate(input: string, start: number, closingTag: string | undefined): ParseResult {
    const nodes: TemplateNode[] = [];
    let cursor = start;

    while (cursor < input.length) {
      const open = input.indexOf('{{', cursor);

      if (open === -1) {
        if (closingTag) {
          throw new Error(`Missing closing tag: {{/${closingTag}}}`);
        }
        this.pushText(nodes, input.slice(cursor), cursor);
        return { nodes, nextIndex: input.length };
      }

      if (open > cursor) {
        this.pushText(nodes, input.slice(cursor, open), cursor);
      }

      const close = input.indexOf('}}', open + 2);
      if (close === -1) {
        throw new Error('Unclosed template expression.');
      }

      const expression = input.slice(open + 2, close).trim();
      if (expression.startsWith('/')) {
        const tagName = expression.slice(1).trim();
        if (closingTag && tagName === closingTag) {
          return { nodes, nextIndex: close + 2 };
        }
        throw new Error(`Unexpected closing tag: {{/${tagName}}}`);
      }

      const control = this.parseControl(expression);
      if (control) {
        const childResult = this.parseTemplate(input, close + 2, control.closingTag);
        nodes.push(this.createControlNode(control, childResult.nodes, open, childResult.nextIndex));
        cursor = childResult.nextIndex;
        continue;
      }

      nodes.push(this.parseExpression(expression, open, close + 2));
      cursor = close + 2;
    }

    if (closingTag) {
      throw new Error(`Missing closing tag: {{/${closingTag}}}`);
    }

    return { nodes, nextIndex: cursor };
  }

  private pushText(nodes: TemplateNode[], value: string, start: number): void {
    if (!value) {
      return;
    }

    const node: TextNode = {
      kind: 'TextNode',
      id: this.createId('text', start),
      value,
      source: {
        start,
        end: start + value.length,
      },
    };
    nodes.push(node);
  }

  private parseControl(expression: string): ParsedControl | undefined {
    const [keyword, ...rest] = expression.split(/\s+/);
    const rawPath = rest.join(' ').trim();

    if (keyword === '#each') {
      return { kind: 'each', path: rawPath, closingTag: 'each' };
    }

    if (keyword === '#each-col') {
      const parsed = this.parsePathWithAttributes(rawPath);
      return { kind: 'each-col', path: parsed.path, attributes: parsed.attributes, closingTag: 'each-col' };
    }

    if (keyword === '#block') {
      return { kind: 'block', path: rawPath, closingTag: 'block' };
    }

    if (keyword === '#grid') {
      return { kind: 'grid', path: rawPath, closingTag: 'grid' };
    }

    if (keyword === '#if') {
      return { kind: 'if', path: rawPath, closingTag: 'if' };
    }

    return undefined;
  }

  private createControlNode(
    control: ParsedControl,
    children: readonly TemplateNode[],
    start: number,
    end: number,
  ): TemplateNode {
    if (!control.path) {
      throw new Error(`Control tag requires a path: ${control.kind}`);
    }

    if (control.kind === 'each') {
      const node: EachNode = {
        kind: 'EachNode',
        id: this.createId('each', start),
        path: control.path,
        children,
        source: { start, end },
      };
      return node;
    }

    if (control.kind === 'each-col') {
      const node: EachColumnNode = {
        kind: 'EachColumnNode',
        id: this.createId('each-col', start),
        path: control.path,
        ...(control.attributes?.span ? { spanPath: control.attributes.span } : {}),
        ...(control.attributes?.render ? { renderPath: control.attributes.render } : {}),
        ...this.parseRowSpanAttribute(control.attributes?.rowspan),
        ...this.parseReserveAttribute(control.attributes?.reserve),
        children,
        source: { start, end },
      };
      return node;
    }

    if (control.kind === 'block') {
      const node: BlockNode = {
        kind: 'BlockNode',
        id: this.createId('block', start),
        path: control.path,
        children,
        source: { start, end },
      };
      return node;
    }

    if (control.kind === 'grid') {
      const [rowPath, columnPath, ...rest] = control.path.split(/\s+/).filter(Boolean);
      if (!rowPath || !columnPath || rest.length > 0) {
        throw new Error('Grid tag requires exactly two paths: {{#grid rows columns}}');
      }

      const node: GridNode = {
        kind: 'GridNode',
        id: this.createId('grid', start),
        rowPath,
        columnPath,
        children,
        source: { start, end },
      };
      return node;
    }

    const node: IfNode = {
      kind: 'IfNode',
      id: this.createId('if', start),
      conditionPath: control.path,
      children,
      source: { start, end },
    };
    return node;
  }

  private parseExpression(expression: string, start: number, end: number): TemplateNode {
    if (expression.startsWith('image ')) {
      const path = expression.slice('image '.length).trim();
      if (!path) {
        throw new Error('Image expression requires a path.');
      }

      const node: ImageNode = {
        kind: 'ImageNode',
        id: this.createId('image', start),
        path,
        options: {},
        source: { start, end },
      };
      return node;
    }

    const helper = this.parseHelper(expression, start, end);
    if (helper) {
      return helper;
    }

    const node: PlaceholderNode = {
      kind: 'PlaceholderNode',
      id: this.createId('placeholder', start),
      path: expression,
      source: { start, end },
    };
    return node;
  }

  private parsePathWithAttributes(input: string): { readonly path: string; readonly attributes: Readonly<Record<string, string>> } {
    const tokens = input.split(/\s+/).filter(Boolean);
    const path = tokens[0] ?? '';
    const attributes: Record<string, string> = {};

    for (const token of tokens.slice(1)) {
      const equals = token.indexOf('=');
      if (equals <= 0 || equals === token.length - 1) {
        throw new Error(`Invalid each-col attribute: ${token}`);
      }

      const key = token.slice(0, equals).trim();
      const value = token.slice(equals + 1).trim();
      if (key !== 'span' && key !== 'rowspan' && key !== 'reserve' && key !== 'render') {
        throw new Error(`Unsupported each-col attribute: ${key}`);
      }

      attributes[key] = value;
    }

    return { path, attributes };
  }

  private parsePositiveInteger(value: string, name: string): number {
    const parsed = Number(value);
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new Error(`${name} must be an integer greater than or equal to 1.`);
    }

    return parsed;
  }

  private parseRowSpanAttribute(value: string | undefined): Pick<EachColumnNode, 'rowSpan' | 'rowSpanPath'> {
    if (!value) {
      return {};
    }

    if (/^\d+$/.test(value)) {
      return { rowSpan: this.parsePositiveInteger(value, 'rowspan') };
    }

    return { rowSpanPath: value };
  }

  private parseReserveAttribute(value: string | undefined): Pick<EachColumnNode, 'reservedColumns' | 'reservedColumnsPath'> {
    if (!value) {
      return {};
    }

    if (/^\d+$/.test(value)) {
      return { reservedColumns: this.parsePositiveInteger(value, 'reserve') };
    }

    return { reservedColumnsPath: value };
  }

  private parseHelper(expression: string, start: number, end: number): HelperNode | undefined {
    const openParen = expression.indexOf('(');
    const closeParen = expression.endsWith(')') ? expression.length - 1 : -1;

    if (openParen <= 0 || closeParen <= openParen) {
      return undefined;
    }

    const name = expression.slice(0, openParen).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
      return undefined;
    }

    return {
      kind: 'HelperNode',
      id: this.createId('helper', start),
      name,
      args: this.parseHelperArgs(expression.slice(openParen + 1, closeParen)),
      source: { start, end },
    };
  }

  private parseHelperArgs(input: string): readonly HelperArgument[] {
    if (!input.trim()) {
      return [];
    }

    return input.split(',').map((rawArg) => {
      const value = rawArg.trim();

      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        return { kind: 'literal', value: value.slice(1, -1) };
      }

      if (value === 'true' || value === 'false') {
        return { kind: 'literal', value: value === 'true' };
      }

      if (value === 'null') {
        return { kind: 'literal', value: null };
      }

      const numberValue = Number(value);
      if (value && Number.isFinite(numberValue) && /^-?\d+(?:\.\d+)?$/.test(value)) {
        return { kind: 'literal', value: numberValue };
      }

      return { kind: 'path', value };
    });
  }

  private createId(prefix: string, start: number): string {
    return `${prefix}_${start}`;
  }
}

interface ParseResult {
  readonly nodes: readonly TemplateNode[];
  readonly nextIndex: number;
}

interface ParsedControl {
  readonly kind: 'each' | 'each-col' | 'block' | 'grid' | 'if';
  readonly path: string;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly closingTag: string;
}
