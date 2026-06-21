import { ExpressionEvaluator } from '../../core/evaluator/expression-evaluator.js';
import type { ExpressionEvaluatorOptions } from '../../core/evaluator/expression-evaluator.js';
import type {
  BlockEndNode,
  BlockStartNode,
  CellAST,
  EachColumnNode,
  EachNode,
  GridNode,
  HelperArgument,
  HelperNode,
  IfNode,
  ImageNode,
  PlaceholderNode,
  TemplateNode,
  WorkbookAST,
  RowAST,
  SheetAST,
} from '../../core/ast/nodes.js';
import { GridContext } from '../../core/grid/grid-context.js';
import { RenderCursor } from '../../core/grid/render-cursor.js';
import type { EvaluationContext } from '../../core/evaluator/evaluation-context.js';
import type { DefaultHelperRegistry } from '../../core/evaluator/helper-registry.js';
import type { EngineRenderOptions, HelperContext } from '../engine/types.js';
import type { RenderOperation, RenderPlan } from './render-plan.js';
import { LimitExceededError } from '../../shared/errors/engine-error.js';
import { LoopContext } from '../../core/render/loop-context.js';
import type { PathResolutionContext } from '../../core/evaluator/json-path-resolver.js';

export class RenderPlanner {
  private readonly evaluator = new ExpressionEvaluator();

  constructor(private readonly helpers: DefaultHelperRegistry) {}

  async createPlan(
    ast: WorkbookAST,
    context: EvaluationContext,
    options: EngineRenderOptions = {},
  ): Promise<RenderPlan> {
    this.validateWorkbookLimits(ast, options);

    const operations: RenderOperation[] = [];
    const warnings: string[] = [];
    const plannedRows = new Set<string>();
    const plannedColumns = new Set<string>();

    for (const sheet of ast.sheets) {
      const blockRegions = this.findBlockRegions(sheet);
      const skippedCells = new Set<string>();
      const deferredOperations: RenderOperation[] = [];
      for (const region of blockRegions) {
        region.cellsToSkip.forEach((key) => skippedCells.add(key));
        const blockPlan = await this.planBlockRegion(sheet.name, region, context, options, warnings);
        operations.push(...blockPlan.operations);
        deferredOperations.push(...blockPlan.deferredOperations);
      }

      for (const row of sheet.rows) {
        for (const cell of row.cells) {
          if (skippedCells.has(this.cellKey(cell))) {
            continue;
          }

          if (cell.nodes.length === 0) {
            continue;
          }

          if (this.isEachColumnCell(cell)) {
            operations.push(...await this.planEachColumnCell(sheet.name, cell, cell.nodes[0], context, options, warnings));
            plannedColumns.add(this.expansionKey(sheet.name, cell.address.column, cell.nodes[0].path));
            continue;
          }

          if (this.isEachRowCell(cell)) {
            operations.push(...await this.planEachRowCell(sheet.name, cell, cell.nodes[0], context, options, warnings));
            plannedRows.add(this.expansionKey(sheet.name, cell.address.row, cell.nodes[0].path));
            continue;
          }

          if (this.isGridCell(cell)) {
            operations.push(...await this.planGridCell(sheet.name, cell, cell.nodes[0], context, options, warnings, plannedRows, plannedColumns));
            continue;
          }

          if (this.isImageCell(cell)) {
            operations.push(...this.planImageCell(sheet.name, cell, cell.nodes[0], context, options));
            continue;
          }

          const value = await this.evaluateCell(cell, context, options, warnings);
          operations.push({
            id: `set_${sheet.name}_${cell.address.row}_${cell.address.column}`,
            type: 'SetCellValue',
            sheetName: sheet.name,
            cell: cell.address,
            value,
          });
        }
      }

      operations.push(...this.sortDeferredDeletes(deferredOperations));
      this.validateOperationLimit(operations.length, options);
    }

    this.validateOperationLimit(operations.length, options);

    return {
      operations,
      warnings,
    };
  }

  private async evaluateCell(
    cell: CellAST,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<unknown> {
    if (cell.nodes.length === 1) {
      return this.evaluateNode(cell.nodes[0]!, context, options, warnings);
    }

    const parts = await Promise.all(
      cell.nodes.map((node) => this.evaluateNode(node, context, options, warnings)),
    );

    return parts.map((part) => this.stringify(part)).join('');
  }

  private isEachColumnCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [EachColumnNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'EachColumnNode';
  }

  private isEachRowCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [EachNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'EachNode';
  }

  private isGridCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [GridNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'GridNode';
  }

  private isImageCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [ImageNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'ImageNode';
  }

  private isBlockStartCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [BlockStartNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'BlockStartNode';
  }

  private isBlockEndCell(cell: CellAST): cell is CellAST & { readonly nodes: readonly [BlockEndNode] } {
    return cell.nodes.length === 1 && cell.nodes[0]?.kind === 'BlockEndNode';
  }

  private async planEachColumnCell(
    sheetName: string,
    cell: CellAST,
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<RenderOperation[]> {
    const collection = this.evaluator.evaluate(
      node.path,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );

    if (!Array.isArray(collection)) {
      warnings.push(`EachColumnNode path không phải array: ${node.path}`);
      return [{
        id: `set_${sheetName}_${cell.address.row}_${cell.address.column}`,
        type: 'SetCellValue',
        sheetName,
        cell: cell.address,
        value: '',
      }];
    }

    if (collection.length === 0) {
      return [{
        id: `set_${sheetName}_${cell.address.row}_${cell.address.column}`,
        type: 'SetCellValue',
        sheetName,
        cell: cell.address,
        value: '',
      }];
    }

    const operations: RenderOperation[] = [];
    if (collection.length > 1) {
      operations.push({
        id: `clone_col_${sheetName}_${cell.address.column}_${collection.length - 1}`,
        type: 'CloneColumn',
        sheetName,
        sourceColumn: cell.address.column,
        targetColumn: cell.address.column + 1,
        count: collection.length - 1,
      });
    }

    for (let index = 0; index < collection.length; index += 1) {
      const itemContext = context.child(collection[index], LoopContext.forItem(index, collection.length, context.loop));
      const parts = await Promise.all(
        node.children.map((child) => this.evaluateNode(child, itemContext, options, warnings)),
      );
      operations.push({
        id: `set_each_col_${sheetName}_${cell.address.row}_${cell.address.column + index}`,
        type: 'SetCellValue',
        sheetName,
        cell: {
          ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
          row: cell.address.row,
          column: cell.address.column + index,
        },
        value: parts.map((part) => this.stringify(part)).join(''),
      });
    }

    return operations;
  }

  private async planEachRowCell(
    sheetName: string,
    cell: CellAST,
    node: EachNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<RenderOperation[]> {
    const collection = this.resolveArray(node.path, context, options, warnings, 'EachNode');
    if (!collection || collection.length === 0) {
      return [this.createSetCellOperation(sheetName, cell.address, '')];
    }

    const operations: RenderOperation[] = [];
    if (collection.length > 1) {
      operations.push({
        id: `clone_row_${sheetName}_${cell.address.row}_${collection.length - 1}`,
        type: 'CloneRow',
        sheetName,
        sourceRow: cell.address.row,
        targetRow: cell.address.row + 1,
        count: collection.length - 1,
      });
    }

    for (let index = 0; index < collection.length; index += 1) {
      const itemContext = context.child(collection[index], LoopContext.forItem(index, collection.length, context.loop));
      operations.push(this.createSetCellOperation(
        sheetName,
        {
          ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
          row: cell.address.row + index,
          column: cell.address.column,
        },
        await this.evaluateChildren(node.children, itemContext, options, warnings),
        `set_each_row_${sheetName}_${cell.address.row + index}_${cell.address.column}`,
      ));
    }

    return operations;
  }

  private async planGridCell(
    sheetName: string,
    cell: CellAST,
    node: GridNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
    plannedRows: Set<string>,
    plannedColumns: Set<string>,
  ): Promise<RenderOperation[]> {
    const rows = this.resolveArray(node.rowPath, context, options, warnings, 'GridNode rows');
    const columns = this.resolveArray(node.columnPath, context, options, warnings, 'GridNode columns');
    if (!rows || !columns || rows.length === 0 || columns.length === 0) {
      return [this.createSetCellOperation(sheetName, cell.address, '')];
    }

    const operations: RenderOperation[] = [];
    if (rows.length > 1 && !plannedRows.has(this.expansionKey(sheetName, cell.address.row, node.rowPath))) {
      operations.push({
        id: `grid_clone_row_${sheetName}_${cell.address.row}_${rows.length - 1}`,
        type: 'CloneRow',
        sheetName,
        sourceRow: cell.address.row,
        targetRow: cell.address.row + 1,
        count: rows.length - 1,
      });
    }

    if (columns.length > 1 && !plannedColumns.has(this.expansionKey(sheetName, cell.address.column, node.columnPath))) {
      operations.push({
        id: `grid_clone_col_${sheetName}_${cell.address.column}_${columns.length - 1}`,
        type: 'CloneColumn',
        sheetName,
        sourceColumn: cell.address.column,
        targetColumn: cell.address.column + 1,
        count: columns.length - 1,
      });
    }

    const cursor = new RenderCursor(cell.address);
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
        const gridContext = new GridContext({
          row: rows[rowIndex],
          column: columns[columnIndex],
          rowIndex,
          columnIndex,
        });
        operations.push(this.createSetCellOperation(
          sheetName,
          cursor.at(rowIndex, columnIndex),
          await this.evaluateChildren(node.children, context.child(gridContext.toCurrentScope()), options, warnings),
          `set_grid_${sheetName}_${cell.address.row + rowIndex}_${cell.address.column + columnIndex}`,
        ));
      }
    }

    return operations;
  }

  private planImageCell(
    sheetName: string,
    cell: CellAST,
    node: ImageNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
  ): RenderOperation[] {
    return [
      this.createSetCellOperation(sheetName, cell.address, ''),
      {
        id: `insert_image_${sheetName}_${cell.address.row}_${cell.address.column}`,
        type: 'InsertImage',
        sheetName,
        cell: cell.address,
        source: this.evaluator.evaluate(
          node.path,
          this.toResolutionContext(context),
          {
            ...this.toEvaluatorOptions(options),
          },
        ),
        ...node.options,
      },
    ];
  }


  private async evaluateNode(
    node: TemplateNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<unknown> {
    switch (node.kind) {
      case 'TextNode':
        return node.value;
      case 'PlaceholderNode':
        return this.resolvePlaceholder(node, context, options);
      case 'HelperNode':
        return this.evaluateHelper(node, context, options);
      case 'IfNode':
        return this.evaluateIf(node, context, options, warnings);
      case 'ImageNode':
        warnings.push(`Image chưa được render ở giai đoạn hiện tại: ${node.path}`);
        return '';
      case 'EachNode':
      case 'EachColumnNode':
      case 'BlockNode':
        return this.evaluateRepeated(node.path, node.children, context, options, warnings);
      case 'GridNode':
        return this.evaluateInlineGrid(node, context, options, warnings);
      case 'BlockStartNode':
      case 'BlockEndNode':
        warnings.push(`Node ${node.kind} cần renderer layout nâng cao, hiện được bỏ qua an toàn.`);
        return '';
      default:
        return '';
    }
  }

  private resolvePlaceholder(
    node: PlaceholderNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
  ): unknown {
    const value = this.evaluator.evaluate(
      node.path,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );

    return value;
  }

  private async evaluateHelper(
    node: HelperNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
  ): Promise<unknown> {
    const args = node.args.map((arg) => this.evaluateHelperArg(arg, context, options));
    const helperContext: HelperContext = {
      data: context.root,
      root: context.root,
      current: context.current,
    };

    return this.helpers.invoke(node.name, args, helperContext);
  }

  private evaluateHelperArg(
    arg: HelperArgument,
    context: EvaluationContext,
    options: EngineRenderOptions,
  ): unknown {
    if (arg.kind === 'literal') {
      return arg.value;
    }

    return this.evaluator.evaluate(
      arg.value,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );
  }

  private async evaluateIf(
    node: IfNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<unknown> {
    const condition = this.evaluator.evaluate(
      node.conditionPath,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );

    if (!condition) {
      return '';
    }

    const parts = await Promise.all(
      node.children.map((child) => this.evaluateNode(child, context, options, warnings)),
    );

    return parts.map((part) => this.stringify(part)).join('');
  }

  private async evaluateRepeated(
    path: string,
    children: readonly TemplateNode[],
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<string> {
    const collection = this.resolveArray(path, context, options, warnings, 'Inline repeat');
    if (!collection || collection.length === 0) {
      return '';
    }

    const parts: string[] = [];
    for (let index = 0; index < collection.length; index += 1) {
      const itemContext = context.child(
        collection[index],
        LoopContext.forItem(index, collection.length, context.loop),
      );
      parts.push(await this.evaluateChildren(children, itemContext, options, warnings));
    }

    return parts.join('');
  }

  private async evaluateInlineGrid(
    node: GridNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<string> {
    const rows = this.resolveArray(node.rowPath, context, options, warnings, 'Inline grid rows');
    const columns = this.resolveArray(node.columnPath, context, options, warnings, 'Inline grid columns');
    if (!rows || !columns || rows.length === 0 || columns.length === 0) {
      return '';
    }

    const parts: string[] = [];
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
        const gridContext = new GridContext({
          row: rows[rowIndex],
          column: columns[columnIndex],
          rowIndex,
          columnIndex,
        });
        parts.push(await this.evaluateChildren(
          node.children,
          context.child(gridContext.toCurrentScope()),
          options,
          warnings,
        ));
      }
    }

    return parts.join('');
  }

  private stringify(value: unknown): string {
    if (value == null) {
      return '';
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return String(value);
  }

  private toEvaluatorOptions(options: EngineRenderOptions): ExpressionEvaluatorOptions {
    return options.missingValue ? { missingValue: options.missingValue } : {};
  }

  private toResolutionContext(context: EvaluationContext): PathResolutionContext {
    return context.loop
      ? { root: context.root, current: context.current, loop: context.loop }
      : { root: context.root, current: context.current };
  }

  private validateWorkbookLimits(ast: WorkbookAST, options: EngineRenderOptions): void {
    const limits = options.limits;
    if (!limits) {
      return;
    }

    if (limits.maxWorksheets !== undefined && ast.sheets.length > limits.maxWorksheets) {
      throw new LimitExceededError('maxWorksheets', ast.sheets.length, limits.maxWorksheets);
    }

    for (const sheet of ast.sheets) {
      const maxRow = sheet.rows.reduce((current, row) => Math.max(current, row.rowNumber), 0);
      const maxColumn = sheet.rows.reduce((current, row) => {
        const rowMaxColumn = row.cells.reduce((column, cell) => Math.max(column, cell.address.column), 0);
        return Math.max(current, rowMaxColumn);
      }, 0);

      if (limits.maxRows !== undefined && maxRow > limits.maxRows) {
        throw new LimitExceededError('maxRows', maxRow, limits.maxRows, { sheetName: sheet.name });
      }

      if (limits.maxColumns !== undefined && maxColumn > limits.maxColumns) {
        throw new LimitExceededError('maxColumns', maxColumn, limits.maxColumns, { sheetName: sheet.name });
      }
    }
  }

  private validateOperationLimit(operationCount: number, options: EngineRenderOptions): void {
    const maxOperations = options.limits?.maxOperations;
    if (maxOperations !== undefined && operationCount > maxOperations) {
      throw new LimitExceededError('maxOperations', operationCount, maxOperations);
    }
  }

  private async evaluateChildren(
    children: readonly TemplateNode[],
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<string> {
    const parts = await Promise.all(
      children.map((child) => this.evaluateNode(child, context, options, warnings)),
    );

    return parts.map((part) => this.stringify(part)).join('');
  }

  private resolveArray(
    path: string,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
    nodeName: string,
  ): readonly unknown[] | undefined {
    const value = this.evaluator.evaluate(
      path,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );

    if (!Array.isArray(value)) {
      warnings.push(`${nodeName} path không phải array: ${path}`);
      return undefined;
    }

    return value;
  }

  private createSetCellOperation(
    sheetName: string,
    cell: CellAST['address'],
    value: unknown,
    id = `set_${sheetName}_${cell.row}_${cell.column}`,
  ): RenderOperation {
    return {
      id,
      type: 'SetCellValue',
      sheetName,
      cell,
      value,
    };
  }

  private expansionKey(sheetName: string, index: number, path: string): string {
    return `${sheetName}:${index}:${path}`;
  }

  private findBlockRegions(sheet: SheetAST): readonly BlockRegion[] {
    const regions: BlockRegion[] = [];
    const flatCells = sheet.rows.flatMap((row) => row.cells.map((cell) => ({ row, cell })));

    for (const { cell } of flatCells) {
      if (!this.isBlockStartCell(cell)) {
        continue;
      }

      const end = flatCells.find(({ cell: candidate }) =>
        candidate.address.row > cell.address.row && this.isBlockEndCell(candidate));

      if (!end) {
        throw new Error(`Missing {{/block}} for block: ${cell.nodes[0].path}`);
      }
      if (!this.isBlockEndCell(end.cell)) {
        throw new Error(`Missing {{/block}} for block: ${cell.nodes[0].path}`);
      }

      const bodyRows = sheet.rows.filter((row) =>
        row.rowNumber > cell.address.row && row.rowNumber < end.cell.address.row);
      const bodyCells = bodyRows.flatMap((row) => row.cells);
      if (bodyCells.length === 0) {
        throw new Error(`Block has no template body: ${cell.nodes[0].path}`);
      }

      const minColumn = Math.min(...bodyCells.map((bodyCell) => bodyCell.address.column));
      const maxColumn = Math.max(...bodyCells.map((bodyCell) => bodyCell.address.column));
      const cellsToSkip = [
        this.cellKey(cell),
        this.cellKey(end.cell),
        ...bodyCells.map((bodyCell) => this.cellKey(bodyCell)),
      ];

      regions.push({
        startCell: cell,
        endCell: end.cell,
        bodyRows,
        bodyCells,
        minColumn,
        maxColumn,
        cellsToSkip,
      });
    }

    return regions;
  }

  private async planBlockRegion(
    sheetName: string,
    region: BlockRegion,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<BlockPlanResult> {
    const collection = this.resolveArray(region.startCell.nodes[0].path, context, options, warnings, 'BlockNode');
    if (!collection || collection.length === 0) {
      return {
        operations: [],
        deferredOperations: [{
          id: `delete_empty_block_${sheetName}_${region.startCell.address.row}`,
          type: 'DeleteRows',
          sheetName,
          startRow: region.startCell.address.row,
          count: region.endCell.address.row - region.startCell.address.row + 1,
        }],
      };
    }

    const operations: RenderOperation[] = [];
    const blockHeight = region.bodyRows.length;
    const sourceRange = {
      ...(region.startCell.address.sheetName ? { sheetName: region.startCell.address.sheetName } : {}),
      start: {
        ...(region.startCell.address.sheetName ? { sheetName: region.startCell.address.sheetName } : {}),
        row: region.bodyRows[0]!.rowNumber,
        column: region.minColumn,
      },
      end: {
        ...(region.startCell.address.sheetName ? { sheetName: region.startCell.address.sheetName } : {}),
        row: region.bodyRows[region.bodyRows.length - 1]!.rowNumber,
        column: region.maxColumn,
      },
    };

    if (collection.length > 1) {
      operations.push({
        id: `clone_block_${sheetName}_${region.startCell.address.row}_${collection.length - 1}`,
        type: 'CloneBlock',
        sheetName,
        sourceRange,
        targetTopLeft: {
          ...(region.startCell.address.sheetName ? { sheetName: region.startCell.address.sheetName } : {}),
          row: sourceRange.end.row + 1,
          column: sourceRange.start.column,
        },
        count: collection.length - 1,
        direction: 'down',
      });
    }

    operations.push(this.createSetCellOperation(sheetName, region.startCell.address, ''));

    for (let itemIndex = 0; itemIndex < collection.length; itemIndex += 1) {
      const itemContext = context.child(
        collection[itemIndex],
        LoopContext.forItem(itemIndex, collection.length, context.loop),
      );
      for (const templateCell of region.bodyCells) {
        if (templateCell.nodes.length === 0) {
          continue;
        }

        const rowOffset = templateCell.address.row - sourceRange.start.row;
        const targetCell = {
          ...(templateCell.address.sheetName ? { sheetName: templateCell.address.sheetName } : {}),
          row: sourceRange.start.row + itemIndex * blockHeight + rowOffset,
          column: templateCell.address.column,
        };
        operations.push(this.createSetCellOperation(
          sheetName,
          targetCell,
          await this.evaluateCell(templateCell, itemContext, options, warnings),
          `set_block_${sheetName}_${targetCell.row}_${targetCell.column}`,
        ));
      }
    }

    operations.push(this.createSetCellOperation(sheetName, {
      ...(region.endCell.address.sheetName ? { sheetName: region.endCell.address.sheetName } : {}),
      row: region.endCell.address.row + (collection.length - 1) * blockHeight,
      column: region.endCell.address.column,
    }, ''));

    return {
      operations,
      deferredOperations: [
        {
          id: `delete_block_end_${sheetName}_${region.endCell.address.row}`,
          type: 'DeleteRows',
          sheetName,
          startRow: region.endCell.address.row + (collection.length - 1) * blockHeight,
          count: 1,
        },
        {
          id: `delete_block_start_${sheetName}_${region.startCell.address.row}`,
          type: 'DeleteRows',
          sheetName,
          startRow: region.startCell.address.row,
          count: 1,
        },
      ],
    };
  }

  private sortDeferredDeletes(operations: readonly RenderOperation[]): readonly RenderOperation[] {
    return [...operations].sort((left, right) => {
      if (left.type === 'DeleteRows' && right.type === 'DeleteRows') {
        return right.startRow - left.startRow;
      }

      return 0;
    });
  }

  private cellKey(cell: CellAST): string {
    return `${cell.address.row}:${cell.address.column}`;
  }
}

interface BlockRegion {
  readonly startCell: CellAST & { readonly nodes: readonly [BlockStartNode] };
  readonly endCell: CellAST & { readonly nodes: readonly [BlockEndNode] };
  readonly bodyRows: readonly RowAST[];
  readonly bodyCells: readonly CellAST[];
  readonly minColumn: number;
  readonly maxColumn: number;
  readonly cellsToSkip: readonly string[];
}

interface BlockPlanResult {
  readonly operations: readonly RenderOperation[];
  readonly deferredOperations: readonly RenderOperation[];
}
