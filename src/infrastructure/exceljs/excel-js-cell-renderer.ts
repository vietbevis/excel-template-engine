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

    return String(value);
  }
}
