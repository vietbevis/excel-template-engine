import { stat } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { dirname } from 'node:path';
import type { RenderLimits } from '../../application/engine/types.js';
import type { TemplateInput } from '../../application/engine/types.js';
import type { AssetResolverOptions, WorkbookRenderer, WorkbookTemplateSource } from '../../application/managers/ports.js';
import type { RenderOperation, RenderPlan } from '../../application/planner/render-plan.js';
import type {
  CellTemplateSource,
  RowTemplateSource,
  WorksheetTemplateSource,
} from '../../core/template/workbook-template-source.js';
import { DefaultAssetResolver } from '../assets/default-asset-resolver.js';
import { ExcelJsWorksheetRenderer } from './excel-js-worksheet-renderer.js';
import { LimitExceededError } from '../../shared/errors/engine-error.js';

type ExcelJsLoadInput = Parameters<ExcelJS.Xlsx['load']>[0];

export class ExcelJsWorkbookRenderer implements WorkbookRenderer {
  private workbook: ExcelJS.Workbook | undefined;
  private templateBaseDir = process.cwd();

  constructor(private readonly options: ExcelJsWorkbookRendererOptions = {}) {}

  async load(input: TemplateInput): Promise<WorkbookTemplateSource> {
    this.workbook = new ExcelJS.Workbook();

    if (typeof input === 'string') {
      await this.assertTemplateFileSize(input);
      this.templateBaseDir = dirname(input);
      await this.workbook.xlsx.readFile(input);
    } else {
      this.templateBaseDir = process.cwd();
      const buffer = this.toBuffer(input);
      this.assertTemplateBufferSize(buffer.byteLength);
      await this.workbook.xlsx.load(buffer as unknown as ExcelJsLoadInput);
    }

    return this.scanWorkbook(this.workbook);
  }

  async apply(plan: RenderPlan): Promise<void> {
    const workbook = this.requireWorkbook();
    const operationsBySheet = new Map<string, RenderOperation[]>();

    for (const operation of plan.operations) {
      const sheetOperations = operationsBySheet.get(operation.sheetName);
      if (sheetOperations) {
        sheetOperations.push(operation);
      } else {
        operationsBySheet.set(operation.sheetName, [operation]);
      }
    }

    await Promise.all([...operationsBySheet].map(async ([sheetName, operations]) => {
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        throw new Error(`Worksheet not found: ${sheetName}`);
      }

      const renderer = new ExcelJsWorksheetRenderer(worksheet, new DefaultAssetResolver({
        ...this.options.assetResolver,
        baseDir: this.options.assetResolver?.baseDir ?? this.templateBaseDir,
      }));
      await renderer.applyAll(operations);
    }));
  }

  async write(): Promise<Uint8Array> {
    const buffer = await this.requireWorkbook().xlsx.writeBuffer();
    return new Uint8Array(buffer);
  }

  private scanWorkbook(workbook: ExcelJS.Workbook): WorkbookTemplateSource {
    return {
      sheets: workbook.worksheets.map((worksheet) => this.scanWorksheet(worksheet)),
    };
  }

  private scanWorksheet(worksheet: ExcelJS.Worksheet): WorksheetTemplateSource {
    const rows: RowTemplateSource[] = [];

    worksheet.eachRow((row, rowNumber) => {
      const cells: CellTemplateSource[] = [];

      row.eachCell((cell, columnNumber) => {
        if (cell.isMerged && cell.master !== cell) {
          return;
        }

        const formula = cell.formula || undefined;
        cells.push({
          address: {
            sheetName: worksheet.name,
            row: rowNumber,
            column: columnNumber,
          },
          value: this.getTemplateValue(cell),
          ...(formula ? { formula } : {}),
        });
      });

      rows.push({
        rowNumber,
        height: row.height,
        cells,
      });
    });

    return {
      name: worksheet.name,
      sheetId: String(worksheet.id),
      rows,
      merges: [],
    };
  }

  private getTemplateValue(cell: ExcelJS.Cell): unknown {
    const value = cell.value;

    if (value && typeof value === 'object' && 'text' in value && typeof value.text === 'string') {
      return value.text;
    }

    if (value && typeof value === 'object' && 'richText' in value) {
      return cell.text;
    }

    return value;
  }

  private toBuffer(input: Exclude<TemplateInput, string>): Buffer {
    if (Buffer.isBuffer(input)) {
      return input;
    }

    if (input instanceof Uint8Array) {
      return Buffer.from(input);
    }

    return Buffer.from(input);
  }

  private async assertTemplateFileSize(filePath: string): Promise<void> {
    const maxTemplateBytes = this.options.limits?.maxTemplateBytes;
    if (maxTemplateBytes === undefined) {
      return;
    }

    const fileStat = await stat(filePath);
    if (fileStat.size > maxTemplateBytes) {
      throw new LimitExceededError('maxTemplateBytes', fileStat.size, maxTemplateBytes, { filePath });
    }
  }

  private assertTemplateBufferSize(byteLength: number): void {
    const maxTemplateBytes = this.options.limits?.maxTemplateBytes;
    if (maxTemplateBytes === undefined || byteLength <= maxTemplateBytes) {
      return;
    }

    throw new LimitExceededError('maxTemplateBytes', byteLength, maxTemplateBytes);
  }

  private requireWorkbook(): ExcelJS.Workbook {
    if (!this.workbook) {
      throw new Error('Workbook has not been loaded.');
    }

    return this.workbook;
  }
}

export interface ExcelJsWorkbookRendererOptions {
  readonly assetResolver?: AssetResolverOptions;
  readonly limits?: Pick<RenderLimits, 'maxTemplateBytes'>;
}
