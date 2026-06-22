import type ExcelJS from 'exceljs';
import type { SetCellValueOperation } from '../../application/planner/render-plan.js';

export class ExcelJsCellRenderer {
  constructor(private readonly cell: ExcelJS.Cell) {}

  applySetCellValue(operation: SetCellValueOperation): void {
    this.cell.value = this.toCellValue(operation.value);
  }

  private toCellValue(value: unknown): ExcelJS.CellValue {
    if (value == null) {
      return null;
    }

    if (
      typeof value === 'string'
      || typeof value === 'number'
      || typeof value === 'boolean'
      || value instanceof Date
    ) {
      return value;
    }

    if (this.isFormulaValue(value)) {
      return {
        formula: value.formula.replaceAll('{row}', String(this.cell.row)),
        ...(value.result !== undefined ? { result: value.result } : {}),
      } as ExcelJS.CellValue;
    }

    if (this.isRichTextValue(value)) {
      return value as ExcelJS.CellValue;
    }

    return String(value);
  }

  private isFormulaValue(value: unknown): value is { readonly formula: string; readonly result?: unknown } {
    return typeof value === 'object'
      && value !== null
      && 'formula' in value
      && typeof (value as { readonly formula?: unknown }).formula === 'string';
  }

  private isRichTextValue(value: unknown): value is ExcelJS.CellRichTextValue {
    return typeof value === 'object'
      && value !== null
      && 'richText' in value
      && Array.isArray((value as { readonly richText?: unknown }).richText);
  }
}
