export class ColumnTreeCompiler {
  compile(
    tree: readonly ColumnTreeBand[],
    options: ColumnTreeCompileOptions,
  ): ColumnTreeCompileResult {
    const sourceGroups = this.expandColumnTree(tree);
    const columns: ColumnTreeColumnDefinition[] = [];
    const groups = sourceGroups.map((group) => {
      const groupColumns = group.columns.map((column) => {
        const fullKey = `${group.key}.${column.key}`;
        if (columns.some((existing) => existing.fullKey === fullKey)) {
          throw new Error(`Duplicate column key: ${fullKey}`);
        }

        const definition: ColumnTreeColumnDefinition = {
          ...column,
          band: group.band,
          groupKey: group.key,
          fullKey,
          columnName: toColumnName(options.startColumn + columns.length),
        };
        columns.push(definition);
        return column;
      });

      return {
        key: group.key,
        band: group.band,
        ...(group.bandRowSpan !== undefined ? { bandRowSpan: group.bandRowSpan } : {}),
        renderHeader: (group.bandRowSpan ?? 1) <= 1,
        name: group.name,
        columns: groupColumns,
        colCount: groupColumns.length,
      };
    });

    return {
      bands: this.buildBands(groups),
      groups,
      columns,
      columnCount: columns.length,
    };
  }

  private expandColumnTree(tree: readonly ColumnTreeBand[]): readonly ExpandedColumnTreeGroup[] {
    const concreteGroups = new Map<string, ExpandedColumnTreeGroup>();
    const expanded: ExpandedColumnTreeGroup[] = [];

    for (const band of tree) {
      const children = band.groups ?? (band.columns ? [band] : []);
      for (const group of children) {
        const next = this.expandGroup(group, band.name, band.rowSpan, concreteGroups);
        concreteGroups.set(next.key, next);
        expanded.push(next);
      }
    }

    return expanded;
  }

  private expandGroup(
    group: ColumnTreeGroup,
    bandName: ColumnTreeHeaderValue,
    bandRowSpan: number | undefined,
    concreteGroups: ReadonlyMap<string, ExpandedColumnTreeGroup>,
  ): ExpandedColumnTreeGroup {
    if (group.derive?.type === 'sumSameKey') {
      return this.expandSumSameKeyGroup(group, bandName, bandRowSpan, concreteGroups);
    }

    return {
      key: group.key,
      band: bandName,
      ...(bandRowSpan !== undefined ? { bandRowSpan } : {}),
      name: group.name,
      columns: (group.columns ?? []).map((column) => ({ ...column })),
    };
  }

  private expandSumSameKeyGroup(
    group: ColumnTreeGroup,
    bandName: ColumnTreeHeaderValue,
    bandRowSpan: number | undefined,
    concreteGroups: ReadonlyMap<string, ExpandedColumnTreeGroup>,
  ): ExpandedColumnTreeGroup {
    const sources = group.derive?.from.map((key) => {
      const source = concreteGroups.get(key);
      if (!source) {
        throw new Error(`Derived group ${group.key} references unknown source group: ${key}`);
      }

      return source;
    }) ?? [];
    const columnKeys = group.derive?.columns
      ?? [...new Set(sources.flatMap((source) => source.columns.map((column) => column.key)))];
    const columns: ColumnTreeColumn[] = columnKeys.map((key) => {
      const refs = sources
        .map((source) => source.columns.some((column) => column.key === key) ? `${source.key}.${key}` : undefined)
        .filter((ref): ref is string => ref !== undefined);
      const sourceColumn = sources.flatMap((source) => source.columns).find((column) => column.key === key);

      return {
        key,
        name: sourceColumn?.name ?? key,
        ...(sourceColumn?.format ? { format: sourceColumn.format } : {}),
        formula: refs.length > 0
          ? { type: 'sumRefs', refs }
          : { type: 'excel', template: '0' },
      };
    });

    if (group.derive?.total) {
      const totalFormat = columns.find((column) => column.format)?.format;
      columns.push({
        key: group.derive.total.key,
        name: group.derive.total.name,
        ...(totalFormat ? { format: totalFormat } : {}),
        formula: {
          type: 'sumRefs',
          refs: columns.map((column) => `${group.key}.${column.key}`),
        },
      });
    }

    return {
      key: group.key,
      band: bandName,
      ...(bandRowSpan !== undefined ? { bandRowSpan } : {}),
      name: group.name,
      columns,
    };
  }

  private buildBands(groups: readonly ColumnTreeGroupWithSpan[]): readonly ColumnTreeBandDefinition[] {
    const bands: ColumnTreeBandDefinition[] = [];
    for (const group of groups) {
      const last = bands[bands.length - 1];
      if (last && last.name === group.band) {
        last.colCount += group.colCount;
        last.rowSpan = Math.max(last.rowSpan, group.bandRowSpan ?? 1);
        continue;
      }

      bands.push({ name: group.band, colCount: group.colCount, rowSpan: group.bandRowSpan ?? 1 });
    }

    return bands;
  }
}

export class FormulaTemplateCompiler {
  constructor(
    private readonly columns: readonly ColumnTreeColumnDefinition[],
    private readonly staticColumns: readonly ColumnTreeStaticColumn[] = [],
  ) {}

  compile(formula: ColumnTreeFormula, context: FormulaCompileContext = {}): string {
    if (formula.type === 'excel') {
      return this.compileText(formula.template, context);
    }

    if (formula.type === 'sumRefs') {
      return `SUM(${formula.refs.map((ref) => this.cell(ref, context)).join(',')})`;
    }

    if (formula.type === 'subtract') {
      return `${this.cell(formula.left, context)}-${this.cell(formula.right, context)}`;
    }

    if (formula.type === 'copy') {
      return this.cell(formula.ref, context);
    }

    return `SUM(${this.range(formula.ref, context)})`;
  }

  refs(formula: ColumnTreeFormula): readonly string[] {
    if (formula.type === 'excel') {
      const refs = [...formula.template.matchAll(/\{(?:range:)?([A-Za-z0-9_.]+|current)\}/g)]
        .map((match) => match[1]!)
        .filter((ref) => ref !== 'current' && !/^[A-Z]+$/.test(ref));
      return [...new Set(refs)];
    }

    if (formula.type === 'sumRefs') {
      return formula.refs;
    }

    if (formula.type === 'subtract') {
      return [formula.left, formula.right];
    }

    return [formula.ref];
  }

  private compileText(template: string, context: FormulaCompileContext): string {
    return template.replace(/\{(range:)?([A-Za-z0-9_.]+|current)\}/g, (_match, rangePrefix: string | undefined, ref: string) => {
      return rangePrefix ? this.range(ref, context) : this.cell(ref, context);
    });
  }

  private cell(ref: string, context: FormulaCompileContext): string {
    const effectiveRef = ref === 'current' ? context.currentRef : ref;
    if (!effectiveRef) {
      throw new Error('Formula reference {current} requires currentRef.');
    }

    const metricColumn = this.columns.find((column) => column.fullKey === effectiveRef);
    if (metricColumn) {
      return `${metricColumn.columnName}${context.row ? `$${context.row}` : '{row}'}`;
    }

    const staticColumn = this.staticColumns.find((column) => column.key === effectiveRef);
    if (staticColumn) {
      return `${staticColumn.column}${context.row ? `$${context.row}` : '{row}'}`;
    }

    throw new Error(`Unknown formula reference: ${ref}`);
  }

  private range(ref: string, context: FormulaCompileContext): string {
    if (!context.range) {
      throw new Error(`Formula range reference ${ref} requires range context.`);
    }

    const effectiveRef = ref === 'current' ? context.currentRef : ref;
    if (!effectiveRef) {
      throw new Error('Formula reference {range:current} requires currentRef.');
    }

    const metricColumn = this.columns.find((column) => column.fullKey === effectiveRef);
    if (metricColumn) {
      return `${metricColumn.columnName}$${context.range.startRow}:${metricColumn.columnName}$${context.range.endRow}`;
    }

    const staticColumn = this.staticColumns.find((column) => column.key === effectiveRef);
    if (staticColumn) {
      return `${staticColumn.column}$${context.range.startRow}:${staticColumn.column}$${context.range.endRow}`;
    }

    throw new Error(`Unknown formula range reference: ${ref}`);
  }
}

function toColumnName(columnNumber: number): string {
  let current = columnNumber;
  let name = '';
  while (current > 0) {
    current -= 1;
    name = String.fromCharCode(65 + (current % 26)) + name;
    current = Math.floor(current / 26);
  }

  return name;
}

export interface ColumnTreeCompileOptions {
  readonly startColumn: number;
}

export interface ColumnTreeCompileResult {
  readonly bands: readonly ColumnTreeBandDefinition[];
  readonly groups: readonly ColumnTreeGroupWithSpan[];
  readonly columns: readonly ColumnTreeColumnDefinition[];
  readonly columnCount: number;
}

export interface ColumnTreeBand {
  readonly key: string;
  readonly name: ColumnTreeHeaderValue;
  readonly rowSpan?: number;
  readonly groups?: readonly ColumnTreeGroup[];
  readonly columns?: readonly ColumnTreeColumn[];
}

export interface ColumnTreeGroup {
  readonly key: string;
  readonly name: ColumnTreeHeaderValue;
  readonly columns?: readonly ColumnTreeColumn[];
  readonly derive?: SumSameKeyDerive;
}

export interface SumSameKeyDerive {
  readonly type: 'sumSameKey';
  readonly from: readonly string[];
  readonly columns?: readonly string[];
  readonly total?: {
    readonly key: string;
    readonly name: string;
  };
}

export interface ColumnTreeColumn {
  readonly key: string;
  readonly name: ColumnTreeHeaderValue;
  readonly formula?: ColumnTreeFormula;
  readonly format?: ColumnTreeCellFormat;
}

export interface ColumnTreeCellFormat {
  readonly numFmt?: string;
  readonly wrapText?: boolean;
}

export type ColumnTreeHeaderValue =
  | string
  | {
    readonly value: string;
    readonly wrapText?: boolean;
    readonly choices?: readonly (string | number | boolean)[];
    readonly choice?: {
      readonly values?: readonly (string | number | boolean)[];
      readonly formula?: string;
      readonly allowBlank?: boolean;
      readonly showErrorMessage?: boolean;
      readonly errorTitle?: string;
      readonly error?: string;
      readonly showInputMessage?: boolean;
      readonly promptTitle?: string;
      readonly prompt?: string;
    };
  };

export type ColumnTreeFormula =
  | { readonly type: 'excel'; readonly template: string }
  | { readonly type: 'sumRefs'; readonly refs: readonly string[] }
  | { readonly type: 'sumRange'; readonly ref: string }
  | { readonly type: 'subtract'; readonly left: string; readonly right: string }
  | { readonly type: 'copy'; readonly ref: string };

export interface ColumnTreeStaticColumn {
  readonly key: string;
  readonly column: string;
}

export interface FormulaCompileContext {
  readonly row?: number;
  readonly currentRef?: string;
  readonly range?: {
    readonly startRow: number;
    readonly endRow: number;
  };
}

export interface ColumnTreeBandDefinition {
  readonly name: ColumnTreeHeaderValue;
  colCount: number;
  rowSpan: number;
}

export interface ColumnTreeGroupWithSpan {
  readonly key: string;
  readonly band: ColumnTreeHeaderValue;
  readonly bandRowSpan?: number;
  readonly renderHeader: boolean;
  readonly name: ColumnTreeHeaderValue;
  readonly columns: readonly ColumnTreeColumn[];
  readonly colCount: number;
}

export interface ColumnTreeColumnDefinition extends ColumnTreeColumn {
  readonly band: ColumnTreeHeaderValue;
  readonly groupKey: string;
  readonly fullKey: string;
  readonly columnName: string;
}

interface ExpandedColumnTreeGroup {
  readonly key: string;
  readonly band: ColumnTreeHeaderValue;
  readonly bandRowSpan?: number;
  readonly name: ColumnTreeHeaderValue;
  readonly columns: readonly ColumnTreeColumn[];
}
