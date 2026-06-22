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
import type { CellDataValidation, CellFormat, RenderOperation, RenderPlan } from './render-plan.js';
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
    sheetContexts: ReadonlyMap<string, EvaluationContext> = new Map(),
  ): Promise<RenderPlan> {
    this.validateWorkbookLimits(ast, options);

    const operations: RenderOperation[] = [];
    const warnings: string[] = [];
    const plannedRows = new Set<string>();
    const plannedColumns = new Set<string>();
    const plannedColumnWidths = new Map<string, number>();

    for (const sheet of ast.sheets) {
      const sheetContext = sheetContexts.get(sheet.name) ?? context;
      const blockRegions = this.findBlockRegions(sheet, sheetContext, options, warnings);
      const skippedCells = new Set<string>();
      const deferredOperations: RenderOperation[] = [];
      const blockRowShifts: BlockRowShift[] = [];
      for (const region of blockRegions) {
        region.cellsToSkip.forEach((key) => skippedCells.add(key));
        const blockPlan = await this.planBlockRegion(sheet.name, region, sheetContext, options, warnings);
        operations.push(...blockPlan.operations);
        deferredOperations.push(...blockPlan.deferredOperations);
        if (blockPlan.rowShiftForFollowingCells > 0) {
          blockRowShifts.push({
            afterRow: region.endCell.address.row,
            rowShift: blockPlan.rowShiftForFollowingCells,
          });
        }
      }

      for (const row of sheet.rows) {
        let columnShiftForRow = 0;
        for (const cell of row.cells) {
          if (skippedCells.has(this.cellKey(cell))) {
            continue;
          }

          if (cell.nodes.length === 0) {
            continue;
          }

          const renderCell = this.shiftCellForPriorColumns(
            this.shiftCellForPriorBlocks(cell, blockRowShifts),
            columnShiftForRow,
          );

          if (this.isEachColumnCell(cell)) {
            operations.push(...await this.planEachColumnCell(
              sheet.name,
              renderCell,
              cell.nodes[0],
              sheetContext,
              options,
              warnings,
              { plannedColumnWidths },
            ));
            columnShiftForRow += Math.max(
              this.resolveEachColumnWidth(cell.nodes[0], sheetContext, options, warnings) - 1,
              0,
            );
            plannedColumns.add(this.expansionKey(sheet.name, cell.address.column, cell.nodes[0].path));
            continue;
          }

          if (this.isEachRowCell(cell)) {
            operations.push(...await this.planEachRowCell(sheet.name, renderCell, cell.nodes[0], sheetContext, options, warnings));
            plannedRows.add(this.expansionKey(sheet.name, cell.address.row, cell.nodes[0].path));
            continue;
          }

          if (this.isGridCell(cell)) {
            operations.push(...await this.planGridCell(sheet.name, renderCell, cell.nodes[0], sheetContext, options, warnings, plannedRows, plannedColumns));
            continue;
          }

          if (this.isImageCell(cell)) {
            operations.push(...this.planImageCell(sheet.name, renderCell, cell.nodes[0], sheetContext, options));
            continue;
          }

          const value = await this.evaluateCell(cell, sheetContext, options, warnings);
          operations.push(this.createSetCellOperation(
            sheet.name,
            renderCell.address,
            value,
            `set_${sheet.name}_${renderCell.address.row}_${renderCell.address.column}`,
            renderCell.address,
          ));
        }
      }

      operations.push(...this.sortDeferredDeletes(deferredOperations));
      if (deferredOperations.length > 0) {
        operations.push({
          id: `cleanup_template_markers_${sheet.name}`,
          type: 'CleanupTemplateMarkers',
          sheetName: sheet.name,
        });
      }
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
    plannerOptions: EachColumnPlannerOptions = {},
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
      return [this.createSetCellOperation(
        sheetName,
        cell.address,
        '',
        `set_${sheetName}_${cell.address.row}_${cell.address.column}`,
        plannerOptions.styleSource ?? cell.address,
      )];
    }

    if (collection.length === 0) {
      return [this.createSetCellOperation(
        sheetName,
        cell.address,
        '',
        `set_${sheetName}_${cell.address.row}_${cell.address.column}`,
        plannerOptions.styleSource ?? cell.address,
      )];
    }

    const items = this.toEachColumnItems(collection, node, context, options, warnings);
    if (items.length === 0) {
      return [this.createSetCellOperation(
        sheetName,
        cell.address,
        '',
        `set_${sheetName}_${cell.address.row}_${cell.address.column}`,
        plannerOptions.styleSource ?? cell.address,
      )];
    }

    const totalColumns = items.reduce((sum, item) => sum + item.span, 0);
    const operations: RenderOperation[] = [];
    const widthKey = this.columnWidthKey(sheetName, cell.address.column);
    const reservedColumns = this.resolveReservedColumns(node, context, options, warnings);
    const existingWidth = Math.max(
      plannerOptions.plannedColumnWidths?.get(widthKey) ?? 1,
      reservedColumns,
    );
    const shouldCloneColumns = plannerOptions.allowClone !== false && totalColumns > existingWidth;
    if (shouldCloneColumns) {
      operations.push({
        id: `clone_col_${sheetName}_${cell.address.column}_${totalColumns - existingWidth}`,
        type: 'CloneColumn',
        sheetName,
        sourceColumn: cell.address.column,
        targetColumn: cell.address.column + existingWidth,
        count: totalColumns - existingWidth,
      });
    }
    plannerOptions.plannedColumnWidths?.set(widthKey, Math.max(existingWidth, totalColumns));

    let columnOffset = 0;
    for (const item of items) {
      const targetColumn = cell.address.column + columnOffset;
      if (!item.render) {
        columnOffset += item.span;
        continue;
      }

      const itemContext = context.child(item.value, LoopContext.forItem(item.index, collection.length, context.loop));
      const value = node.children.length === 1
        ? await this.evaluateNode(node.children[0]!, itemContext, options, warnings)
        : (await Promise.all(
          node.children.map((child) => this.evaluateNode(child, itemContext, options, warnings)),
        )).map((part) => this.stringify(part)).join('');
      operations.push(this.createSetCellOperation(
        sheetName,
        {
          ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
          row: cell.address.row,
          column: targetColumn,
        },
        value,
        `set_each_col_${sheetName}_${cell.address.row}_${targetColumn}`,
        plannerOptions.styleSource ?? cell.address,
      ));

      if (item.span > 1 || item.rowSpan > 1) {
        operations.push({
          id: `merge_each_col_${sheetName}_${cell.address.row}_${targetColumn}`,
          type: 'ApplyMerge',
          sheetName,
          range: {
            ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
            start: {
              ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
              row: cell.address.row,
              column: targetColumn,
            },
            end: {
              ...(cell.address.sheetName ? { sheetName: cell.address.sheetName } : {}),
              row: cell.address.row + item.rowSpan - 1,
              column: targetColumn + item.span - 1,
            },
          },
        });
      }

      columnOffset += item.span;
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
        cell.address,
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
          cell.address,
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
    styleSource?: CellAST['address'],
  ): RenderOperation {
    const cellValue = this.normalizeCellRenderValue(value);
    return {
      id,
      type: 'SetCellValue',
      sheetName,
      cell,
      value: cellValue.value,
      ...(styleSource && !this.isSameAddress(styleSource, cell) ? { styleSource } : {}),
      ...(cellValue.format ? { format: cellValue.format } : {}),
      ...(cellValue.dataValidation ? { dataValidation: cellValue.dataValidation } : {}),
    };
  }

  private normalizeCellRenderValue(value: unknown): NormalizedCellRenderValue {
    if (!this.isCellRenderValue(value)) {
      return { value };
    }

    const format: { wrapText?: boolean; numFmt?: string } = {};
    if (typeof value.wrapText === 'boolean') {
      format.wrapText = value.wrapText;
    }
    if (typeof value.numFmt === 'string') {
      format.numFmt = value.numFmt;
    }

    const dataValidation = this.normalizeChoiceValidation(value);

    return {
      value: value.value,
      ...(Object.keys(format).length > 0 ? { format } : {}),
      ...(dataValidation ? { dataValidation } : {}),
    };
  }

  private isCellRenderValue(value: unknown): value is CellRenderValueLike {
    return typeof value === 'object'
      && value !== null
      && 'value' in value
      && ('wrapText' in value || 'numFmt' in value || 'choices' in value || 'choice' in value);
  }

  private normalizeChoiceValidation(value: CellRenderValueLike): CellDataValidation | undefined {
    const rawChoice = value.choice ?? value.choices;
    if (!rawChoice) {
      return undefined;
    }

    if (Array.isArray(rawChoice)) {
      return {
        type: 'choice',
        values: rawChoice.filter(this.isChoiceValue),
      };
    }

    if (typeof rawChoice !== 'object') {
      return undefined;
    }

    const choice = rawChoice as ChoiceValidationLike;
    const values = Array.isArray(choice.values)
      ? choice.values.filter(this.isChoiceValue)
      : undefined;
    const formula = typeof choice.formula === 'string' && choice.formula.trim() !== ''
      ? choice.formula
      : undefined;

    if (!values && !formula) {
      return undefined;
    }

    return {
      type: 'choice',
      ...(values ? { values } : {}),
      ...(formula ? { formula } : {}),
      ...(typeof choice.allowBlank === 'boolean' ? { allowBlank: choice.allowBlank } : {}),
      ...(typeof choice.showErrorMessage === 'boolean' ? { showErrorMessage: choice.showErrorMessage } : {}),
      ...(typeof choice.errorTitle === 'string' ? { errorTitle: choice.errorTitle } : {}),
      ...(typeof choice.error === 'string' ? { error: choice.error } : {}),
      ...(typeof choice.showInputMessage === 'boolean' ? { showInputMessage: choice.showInputMessage } : {}),
      ...(typeof choice.promptTitle === 'string' ? { promptTitle: choice.promptTitle } : {}),
      ...(typeof choice.prompt === 'string' ? { prompt: choice.prompt } : {}),
    };
  }

  private isChoiceValue(value: unknown): value is string | number | boolean {
    return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
  }

  private isSameAddress(left: CellAST['address'], right: CellAST['address']): boolean {
    return left.sheetName === right.sheetName
      && left.row === right.row
      && left.column === right.column;
  }

  private expansionKey(sheetName: string, index: number, path: string): string {
    return `${sheetName}:${index}:${path}`;
  }

  private columnWidthKey(sheetName: string, column: number): string {
    return `${sheetName}:${column}`;
  }

  private findBlockRegions(
    sheet: SheetAST,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): readonly BlockRegion[] {
    const regions: BlockRegion[] = [];
    const flatCells = sheet.rows.flatMap((row) => row.cells.map((cell) => ({ row, cell })));
    const stack: BlockRegionFrame[] = [];

    for (const { cell } of flatCells) {
      if (this.isBlockStartCell(cell)) {
        stack.push({
          startCell: cell,
          childRegions: [],
        });
        continue;
      }

      if (!this.isBlockEndCell(cell)) {
        continue;
      }

      const frame = stack.pop();
      if (!frame) {
        throw new Error('Unexpected {{/block}}.');
      }

      const bodyRows = sheet.rows.filter((row) =>
        row.rowNumber > frame.startCell.address.row && row.rowNumber < cell.address.row);
      const bodyCells = bodyRows.flatMap((row) => row.cells);
      if (bodyCells.length === 0) {
        throw new Error(`Block has no template body: ${frame.startCell.nodes[0].path}`);
      }

      const minColumn = Math.min(...bodyCells.map((bodyCell) => bodyCell.address.column));
      const maxColumn = Math.max(...bodyCells.map((bodyCell) => this.cellMaxColumn(bodyCell, context, options, warnings)));
      const cellsToSkip = [
        this.cellKey(frame.startCell),
        this.cellKey(cell),
        ...bodyCells.map((bodyCell) => this.cellKey(bodyCell)),
      ];

      const region: BlockRegion = {
        startCell: frame.startCell,
        endCell: cell,
        bodyRows,
        bodyCells,
        minColumn,
        maxColumn,
        cellsToSkip,
        childRegions: frame.childRegions,
      };

      const parent = stack[stack.length - 1];
      if (parent) {
        parent.childRegions.push(region);
      } else {
        regions.push(region);
      }
    }

    if (stack.length > 0) {
      throw new Error(`Missing {{/block}} for block: ${stack[stack.length - 1]!.startCell.nodes[0].path}`);
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
    if (region.childRegions.length > 0) {
      return this.planNestedBlockRegion(sheetName, region, context, options, warnings);
    }

    const collection = this.resolveArray(region.startCell.nodes[0].path, context, options, warnings, 'BlockNode');
    if (!collection || collection.length === 0) {
      return {
        operations: [],
        rowShiftForFollowingCells: 0,
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

        if (this.isEachColumnCell(templateCell)) {
          operations.push(...await this.planEachColumnCell(
            sheetName,
            {
              ...templateCell,
              address: targetCell,
            },
            templateCell.nodes[0],
            itemContext,
            options,
            warnings,
            { allowClone: false, styleSource: targetCell },
          ));
          continue;
        }

        operations.push(this.createSetCellOperation(
          sheetName,
          targetCell,
          await this.evaluateCell(templateCell, itemContext, options, warnings),
          `set_block_${sheetName}_${targetCell.row}_${targetCell.column}`,
          targetCell,
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
      rowShiftForFollowingCells: Math.max(collection.length - 1, 0) * blockHeight,
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

  private async planNestedBlockRegion(
    sheetName: string,
    region: BlockRegion,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<BlockPlanResult> {
    const childRegion = region.childRegions[0];
    if (!childRegion || region.childRegions.length > 1 || childRegion.bodyRows.length !== 1) {
      warnings.push('Nested BlockNode hiện hỗ trợ một inner block với một body row.');
      return this.planFlatBlockRegion(sheetName, region, context, options, warnings);
    }

    const collection = this.resolveArray(region.startCell.nodes[0].path, context, options, warnings, 'BlockNode');
    if (!collection || collection.length === 0) {
      return {
        operations: [],
        rowShiftForFollowingCells: 0,
        deferredOperations: [{
          id: `delete_empty_nested_block_${sheetName}_${region.startCell.address.row}`,
          type: 'DeleteRows',
          sheetName,
          startRow: region.startCell.address.row,
          count: region.endCell.address.row - region.startCell.address.row + 1,
        }],
      };
    }

    const blockHeight = region.bodyRows.length;
    const sourceRange = this.createBlockSourceRange(region);
    const operations: RenderOperation[] = [];
    const innerCollections = collection.map((item, itemIndex) => {
      const itemContext = context.child(item, LoopContext.forItem(itemIndex, collection.length, context.loop));
      return this.resolveArray(childRegion.startCell.nodes[0].path, itemContext, options, warnings, 'Nested BlockNode') ?? [];
    });

    if (collection.length > 1) {
      operations.push({
        id: `clone_nested_block_${sheetName}_${region.startCell.address.row}_${collection.length - 1}`,
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

    for (let itemIndex = collection.length - 1; itemIndex >= 0; itemIndex -= 1) {
      const innerCollection = innerCollections[itemIndex] ?? [];
      if (innerCollection.length <= 1) {
        continue;
      }

      operations.push({
        id: `clone_nested_inner_row_${sheetName}_${itemIndex}_${innerCollection.length - 1}`,
        type: 'CloneRow',
        sheetName,
        sourceRow: childRegion.bodyRows[0]!.rowNumber + itemIndex * blockHeight,
        targetRow: childRegion.bodyRows[0]!.rowNumber + itemIndex * blockHeight + 1,
        count: innerCollection.length - 1,
      });
    }

    const templateRowsToSkip = new Set<number>();
    for (let row = childRegion.startCell.address.row; row <= childRegion.endCell.address.row; row += 1) {
      templateRowsToSkip.add(row);
    }

    for (let itemIndex = 0; itemIndex < collection.length; itemIndex += 1) {
      const extraRowsBefore = this.countNestedExtraRows(innerCollections, itemIndex);
      const itemContext = context.child(
        collection[itemIndex],
        LoopContext.forItem(itemIndex, collection.length, context.loop),
      );

      for (const templateCell of region.bodyCells) {
        if (templateRowsToSkip.has(templateCell.address.row) || templateCell.nodes.length === 0) {
          continue;
        }

        const rowOffset = templateCell.address.row - sourceRange.start.row;
        const rowShiftWithinItem = templateCell.address.row > childRegion.endCell.address.row
          ? Math.max((innerCollections[itemIndex]?.length ?? 0) - 1, 0)
          : 0;
        const targetCell = {
          ...(templateCell.address.sheetName ? { sheetName: templateCell.address.sheetName } : {}),
          row: sourceRange.start.row + itemIndex * blockHeight + extraRowsBefore + rowOffset + rowShiftWithinItem,
          column: templateCell.address.column,
        };

        if (this.isEachColumnCell(templateCell)) {
          operations.push(...await this.planEachColumnCell(
            sheetName,
            {
              ...templateCell,
              address: targetCell,
            },
            templateCell.nodes[0],
            itemContext,
            options,
            warnings,
            { allowClone: false, styleSource: targetCell },
          ));
          continue;
        }

        operations.push(this.createSetCellOperation(
          sheetName,
          targetCell,
          await this.evaluateCell(templateCell, itemContext, options, warnings),
          `set_nested_block_${sheetName}_${targetCell.row}_${targetCell.column}`,
          targetCell,
        ));
        if (templateCell.sourceRange) {
          operations.push(this.createMergeFromTemplateRange(
            sheetName,
            templateCell.sourceRange,
            targetCell.row - templateCell.sourceRange.start.row,
            targetCell.column - templateCell.sourceRange.start.column,
            `merge_nested_block_${sheetName}_${targetCell.row}_${targetCell.column}`,
          ));
        }
      }

      const innerCollection = innerCollections[itemIndex] ?? [];
      for (let innerIndex = 0; innerIndex < innerCollection.length; innerIndex += 1) {
        const innerContext = itemContext.child(
          innerCollection[innerIndex],
          LoopContext.forItem(innerIndex, innerCollection.length, itemContext.loop),
        );
        for (const templateCell of childRegion.bodyCells) {
          if (templateCell.nodes.length === 0) {
            continue;
          }

          const targetCell = {
            ...(templateCell.address.sheetName ? { sheetName: templateCell.address.sheetName } : {}),
            row: childRegion.bodyRows[0]!.rowNumber + itemIndex * blockHeight + extraRowsBefore + innerIndex,
            column: templateCell.address.column,
          };
          if (this.isEachColumnCell(templateCell)) {
            operations.push(...await this.planEachColumnCell(
              sheetName,
              {
                ...templateCell,
                address: targetCell,
              },
              templateCell.nodes[0],
              innerContext,
              options,
              warnings,
              { allowClone: false, styleSource: targetCell },
            ));
            continue;
          }

          operations.push(this.createSetCellOperation(
            sheetName,
            targetCell,
            await this.evaluateCell(templateCell, innerContext, options, warnings),
            `set_nested_inner_block_${sheetName}_${targetCell.row}_${targetCell.column}`,
            targetCell,
          ));
        }
      }
    }

    const totalExtraRows = this.countNestedExtraRows(innerCollections, collection.length);
    const deferredOperations: RenderOperation[] = [
      {
        id: `delete_nested_outer_end_${sheetName}_${region.endCell.address.row}`,
        type: 'DeleteRows',
        sheetName,
        startRow: region.endCell.address.row + (collection.length - 1) * blockHeight + totalExtraRows,
        count: 1,
      },
    ];

    for (let itemIndex = collection.length - 1; itemIndex >= 0; itemIndex -= 1) {
      const extraRowsBefore = this.countNestedExtraRows(innerCollections, itemIndex);
      const innerExtraRows = Math.max((innerCollections[itemIndex]?.length ?? 0) - 1, 0);
      deferredOperations.push(
        {
          id: `delete_nested_inner_end_${sheetName}_${itemIndex}`,
          type: 'DeleteRows',
          sheetName,
          startRow: childRegion.endCell.address.row + itemIndex * blockHeight + extraRowsBefore + innerExtraRows,
          count: 1,
        },
        {
          id: `delete_nested_inner_start_${sheetName}_${itemIndex}`,
          type: 'DeleteRows',
          sheetName,
          startRow: childRegion.startCell.address.row + itemIndex * blockHeight + extraRowsBefore,
          count: 1,
        },
      );
    }

    deferredOperations.push({
      id: `delete_nested_outer_start_${sheetName}_${region.startCell.address.row}`,
      type: 'DeleteRows',
      sheetName,
      startRow: region.startCell.address.row,
      count: 1,
    });

    return {
      operations,
      rowShiftForFollowingCells: Math.max(collection.length - 1, 0) * blockHeight + totalExtraRows,
      deferredOperations,
    };
  }

  private async planFlatBlockRegion(
    sheetName: string,
    region: BlockRegion,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): Promise<BlockPlanResult> {
    const { childRegions: _childRegions, ...flatRegion } = region;
    return this.planBlockRegion(
      sheetName,
      {
        ...flatRegion,
        childRegions: [],
      },
      context,
      options,
      warnings,
    );
  }

  private createBlockSourceRange(region: BlockRegion) {
    return {
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
  }

  private createMergeFromTemplateRange(
    sheetName: string,
    sourceRange: NonNullable<CellAST['sourceRange']>,
    rowDelta: number,
    columnDelta: number,
    id: string,
  ): RenderOperation {
    return {
      id,
      type: 'ApplyMerge',
      sheetName,
      range: {
        ...(sourceRange.sheetName ? { sheetName: sourceRange.sheetName } : {}),
        start: {
          ...(sourceRange.start.sheetName ? { sheetName: sourceRange.start.sheetName } : {}),
          row: sourceRange.start.row + rowDelta,
          column: sourceRange.start.column + columnDelta,
        },
        end: {
          ...(sourceRange.end.sheetName ? { sheetName: sourceRange.end.sheetName } : {}),
          row: sourceRange.end.row + rowDelta,
          column: sourceRange.end.column + columnDelta,
        },
      },
    };
  }

  private countNestedExtraRows(collections: readonly (readonly unknown[])[], endExclusive: number): number {
    let total = 0;
    for (let index = 0; index < endExclusive; index += 1) {
      total += Math.max((collections[index]?.length ?? 0) - 1, 0);
    }

    return total;
  }

  private shiftCellForPriorBlocks(cell: CellAST, shifts: readonly BlockRowShift[]): CellAST {
    const rowShift = shifts
      .filter((shift) => cell.address.row > shift.afterRow)
      .reduce((sum, shift) => sum + shift.rowShift, 0);
    if (rowShift === 0) {
      return cell;
    }

    return {
      ...cell,
      address: {
        ...cell.address,
        row: cell.address.row + rowShift,
      },
    };
  }

  private shiftCellForPriorColumns(cell: CellAST, columnShift: number): CellAST {
    if (columnShift === 0) {
      return cell;
    }

    return {
      ...cell,
      address: {
        ...cell.address,
        column: cell.address.column + columnShift,
      },
    };
  }

  private resolveEachColumnWidth(
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): number {
    const collection = this.evaluator.evaluate(
      node.path,
      this.toResolutionContext(context),
      {
        ...this.toEvaluatorOptions(options),
      },
    );
    if (!Array.isArray(collection) || collection.length === 0) {
      return this.resolveReservedColumns(node, context, options, warnings);
    }

    const totalColumns = this.toEachColumnItems(collection, node, context, options, warnings)
      .reduce((sum, item) => sum + item.span, 0);
    return Math.max(totalColumns, this.resolveReservedColumns(node, context, options, warnings), 1);
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

  private cellMaxColumn(
    cell: CellAST,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): number {
    const eachColumnNode = cell.nodes.length === 1 && cell.nodes[0]?.kind === 'EachColumnNode'
      ? cell.nodes[0]
      : undefined;
    if (!eachColumnNode) {
      return cell.address.column;
    }

    return cell.address.column + this.resolveReservedColumns(eachColumnNode, context, options, warnings) - 1;
  }

  private toEachColumnItems(
    collection: readonly unknown[],
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): readonly EachColumnItem[] {
    const items: EachColumnItem[] = [];
    for (let index = 0; index < collection.length; index += 1) {
      const value = collection[index];
      const span = this.resolveColumnSpan(value, node, context, options, warnings);
      if (span <= 0) {
        continue;
      }

      items.push({
        index,
        value,
        span,
        render: this.resolveColumnRender(value, node, context, options, warnings),
        rowSpan: this.resolveColumnRowSpan(value, node, context, options, warnings),
      });
    }

    return items;
  }

  private resolveColumnSpan(
    value: unknown,
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): number {
    if (!node.spanPath) {
      return 1;
    }

    const rawSpan = this.evaluator.evaluate(
      node.spanPath,
      this.toResolutionContext(context.child(value)),
      {
        missingValue: 'null',
        ...this.toEvaluatorOptions(options),
      },
    );

    if (rawSpan === null || rawSpan === undefined || rawSpan === '') {
      return 1;
    }

    const span = Number(rawSpan);
    if (!Number.isInteger(span)) {
      warnings.push(`EachColumnNode span không phải integer: ${node.spanPath}`);
      return 1;
    }

    return span;
  }

  private resolveReservedColumns(
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): number {
    if (!node.reservedColumnsPath) {
      return node.reservedColumns ?? 1;
    }

    const rawReservedColumns = this.evaluator.evaluate(
      node.reservedColumnsPath,
      this.toResolutionContext(context),
      {
        missingValue: 'null',
        ...this.toEvaluatorOptions(options),
      },
    );

    if (rawReservedColumns === null || rawReservedColumns === undefined || rawReservedColumns === '') {
      return node.reservedColumns ?? 1;
    }

    const reservedColumns = Number(rawReservedColumns);
    if (!Number.isInteger(reservedColumns) || reservedColumns < 1) {
      warnings.push(`EachColumnNode reserve không phải positive integer: ${node.reservedColumnsPath}`);
      return node.reservedColumns ?? 1;
    }

    return reservedColumns;
  }

  private resolveColumnRowSpan(
    value: unknown,
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): number {
    if (!node.rowSpanPath) {
      return node.rowSpan ?? 1;
    }

    const rawRowSpan = this.evaluator.evaluate(
      node.rowSpanPath,
      this.toResolutionContext(context.child(value)),
      {
        missingValue: 'null',
        ...this.toEvaluatorOptions(options),
      },
    );

    if (rawRowSpan === null || rawRowSpan === undefined || rawRowSpan === '') {
      return node.rowSpan ?? 1;
    }

    const rowSpan = Number(rawRowSpan);
    if (!Number.isInteger(rowSpan) || rowSpan < 1) {
      warnings.push(`EachColumnNode rowspan không phải positive integer: ${node.rowSpanPath}`);
      return node.rowSpan ?? 1;
    }

    return rowSpan;
  }

  private resolveColumnRender(
    value: unknown,
    node: EachColumnNode,
    context: EvaluationContext,
    options: EngineRenderOptions,
    warnings: string[],
  ): boolean {
    if (!node.renderPath) {
      return true;
    }

    const rawRender = this.evaluator.evaluate(
      node.renderPath,
      this.toResolutionContext(context.child(value)),
      {
        missingValue: 'null',
        ...this.toEvaluatorOptions(options),
      },
    );

    if (rawRender === null || rawRender === undefined || rawRender === '') {
      return true;
    }

    if (typeof rawRender === 'boolean') {
      return rawRender;
    }

    if (typeof rawRender === 'number') {
      return rawRender !== 0;
    }

    if (typeof rawRender === 'string') {
      const normalized = rawRender.trim().toLowerCase();
      if (normalized === 'false' || normalized === '0' || normalized === 'no') {
        return false;
      }
      if (normalized === 'true' || normalized === '1' || normalized === 'yes') {
        return true;
      }
    }

    warnings.push(`EachColumnNode render không phải boolean-compatible: ${node.renderPath}`);
    return true;
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
  readonly childRegions: readonly BlockRegion[];
}

interface BlockPlanResult {
  readonly operations: readonly RenderOperation[];
  readonly rowShiftForFollowingCells: number;
  readonly deferredOperations: readonly RenderOperation[];
}

interface BlockRowShift {
  readonly afterRow: number;
  readonly rowShift: number;
}

interface BlockRegionFrame {
  readonly startCell: CellAST & { readonly nodes: readonly [BlockStartNode] };
  readonly childRegions: BlockRegion[];
}

interface EachColumnPlannerOptions {
  readonly allowClone?: boolean;
  readonly plannedColumnWidths?: Map<string, number>;
  readonly styleSource?: CellAST['address'];
}

interface EachColumnItem {
  readonly index: number;
  readonly value: unknown;
  readonly span: number;
  readonly render: boolean;
  readonly rowSpan: number;
}

interface NormalizedCellRenderValue {
  readonly value: unknown;
  readonly format?: CellFormat;
  readonly dataValidation?: CellDataValidation;
}

interface CellRenderValueLike {
  readonly value: unknown;
  readonly wrapText?: unknown;
  readonly numFmt?: unknown;
  readonly choices?: unknown;
  readonly choice?: ChoiceValidationLike | readonly unknown[];
}

interface ChoiceValidationLike {
  readonly values?: unknown;
  readonly formula?: unknown;
  readonly allowBlank?: unknown;
  readonly showErrorMessage?: unknown;
  readonly errorTitle?: unknown;
  readonly error?: unknown;
  readonly showInputMessage?: unknown;
  readonly promptTitle?: unknown;
  readonly prompt?: unknown;
}
