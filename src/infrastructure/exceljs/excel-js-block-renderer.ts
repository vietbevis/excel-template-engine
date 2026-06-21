import type ExcelJS from 'exceljs';
import type { CloneBlockOperation } from '../../application/planner/render-plan.js';
import { ExcelJsBlockCloneManager } from './excel-js-block-clone-manager.js';

export class ExcelJsBlockRenderer {
  constructor(private readonly worksheet: ExcelJS.Worksheet) {}

  render(operation: CloneBlockOperation): Promise<void> {
    return new ExcelJsBlockCloneManager(this.worksheet).clone(operation);
  }
}
