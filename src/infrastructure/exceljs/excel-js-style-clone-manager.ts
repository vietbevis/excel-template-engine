import type ExcelJS from 'exceljs';
import type { StyleCloneManager } from '../../application/managers/ports.js';
import type { CellAddress } from '../../shared/address/address.js';

export class ExcelJsStyleCloneManager implements StyleCloneManager {
  constructor(private readonly workbook: ExcelJS.Workbook) {}

  async cloneCellStyle(source: CellAddress, target: CellAddress): Promise<void> {
    const sourceCell = this.getCell(source);
    const targetCell = this.getCell(target);

    targetCell.style = this.cloneStyle(sourceCell.style);
  }

  async cloneRowStyle(sheetName: string, sourceRow: number, targetRow: number): Promise<void> {
    const worksheet = this.getWorksheet(sheetName);
    const source = worksheet.getRow(sourceRow);
    const target = worksheet.getRow(targetRow);

    this.cloneStyleFields(source, target);

    if (source.height !== undefined) {
      target.height = source.height;
    }
  }

  async cloneColumnStyle(sheetName: string, sourceColumn: number, targetColumn: number): Promise<void> {
    const worksheet = this.getWorksheet(sheetName);
    const source = worksheet.getColumn(sourceColumn);
    const target = worksheet.getColumn(targetColumn);

    this.cloneStyleFields(source, target);

    if (source.width !== undefined) {
      target.width = source.width;
    }
  }

  private getCell(address: CellAddress): ExcelJS.Cell {
    if (!address.sheetName) {
      throw new Error('CellAddress.sheetName is required for ExcelJS style cloning.');
    }

    return this.getWorksheet(address.sheetName).getCell(address.row, address.column);
  }

  private getWorksheet(sheetName: string): ExcelJS.Worksheet {
    const worksheet = this.workbook.getWorksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Worksheet not found: ${sheetName}`);
    }

    return worksheet;
  }

  private cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
    return this.deepClone(style);
  }

  private cloneStyleFields(source: Partial<ExcelJS.Style>, target: Partial<ExcelJS.Style>): void {
    this.setStyleField(target, 'numFmt', source.numFmt);
    this.setStyleField(target, 'font', this.deepClone(source.font));
    this.setStyleField(target, 'fill', this.deepClone(source.fill));
    this.setStyleField(target, 'border', this.deepClone(source.border));
    this.setStyleField(target, 'alignment', this.deepClone(source.alignment));
    this.setStyleField(target, 'protection', this.deepClone(source.protection));
  }

  private setStyleField<TKey extends keyof ExcelJS.Style>(
    target: Partial<ExcelJS.Style>,
    key: TKey,
    value: ExcelJS.Style[TKey] | undefined,
  ): void {
    if (value === undefined) {
      delete target[key];
      return;
    }

    target[key] = value;
  }

  private deepClone<T>(value: T): T {
    if (value === undefined || value === null) {
      return value;
    }

    if (typeof structuredClone === 'function') {
      return structuredClone(value);
    }

    return JSON.parse(JSON.stringify(value)) as T;
  }
}
