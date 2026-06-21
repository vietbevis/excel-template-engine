import type ExcelJS from 'exceljs';
import type { SetCellValueOperation } from '../../application/planner/render-plan.js';
import { ExcelJsCellRenderer } from './excel-js-cell-renderer.js';

export class ExcelJsRowRenderer {
  constructor(private readonly row: ExcelJS.Row) {}

  applySetCellValue(operation: SetCellValueOperation): void {
    const cell = this.row.getCell(operation.cell.column);
    new ExcelJsCellRenderer(cell).applySetCellValue(operation);
  }
}
