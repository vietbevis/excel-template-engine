import { stat } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { dirname } from 'node:path';
import type { RenderLimits } from '../../application/engine/types.js';
import type { WorkbookRenderConfig } from '../../application/engine/types.js';
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
import { MergeRange } from '../../core/merge/merge-range.js';

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

  async prepare(config: WorkbookRenderConfig): Promise<WorkbookTemplateSource> {
    const workbook = this.requireWorkbook();
    const worksheets = config.worksheets ?? [];
    const sourceNamesToDelete = new Set<string>();

    for (const worksheetConfig of worksheets) {
      if (worksheetConfig.deleteSource) {
        sourceNamesToDelete.add(worksheetConfig.sourceName);
      }

      if (worksheetConfig.name === worksheetConfig.sourceName) {
        continue;
      }

      if (workbook.getWorksheet(worksheetConfig.name)) {
        throw new Error(`Worksheet already exists: ${worksheetConfig.name}`);
      }

      const source = workbook.getWorksheet(worksheetConfig.sourceName);
      if (!source) {
        throw new Error(`Worksheet template not found: ${worksheetConfig.sourceName}`);
      }

      this.cloneWorksheet(source, worksheetConfig.name);
    }

    for (const sourceName of sourceNamesToDelete) {
      const isStillTarget = worksheets.some((worksheet) => worksheet.name === sourceName);
      if (isStillTarget) {
        continue;
      }

      const worksheet = workbook.getWorksheet(sourceName);
      if (worksheet) {
        workbook.removeWorksheet(worksheet.id);
      }
    }

    return this.scanWorkbook(workbook);
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

      const renderer = new ExcelJsWorksheetRenderer(
        worksheet,
        new DefaultAssetResolver({
          ...this.options.assetResolver,
          baseDir: this.options.assetResolver?.baseDir ?? this.templateBaseDir,
        }),
        { preserveFormulas: this.options.preserveFormulas ?? false },
      );
      await renderer.applyAll(this.sortSheetOperations(operations));
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

  private cloneWorksheet(source: ExcelJS.Worksheet, targetName: string): ExcelJS.Worksheet {
    const workbook = this.requireWorkbook();
    const target = workbook.addWorksheet(targetName, {
      properties: this.deepClone(source.properties),
      pageSetup: this.deepClone(source.pageSetup),
      views: this.deepClone(source.views),
      state: source.state,
    });

    if (source.autoFilter) {
      target.autoFilter = this.deepClone(source.autoFilter);
    }

    for (let columnNumber = 1; columnNumber <= source.columnCount; columnNumber += 1) {
      const sourceColumn = source.getColumn(columnNumber);
      const targetColumn = target.getColumn(columnNumber);
      if (sourceColumn.width !== undefined) {
        targetColumn.width = sourceColumn.width;
      }
      if (sourceColumn.hidden !== undefined) {
        targetColumn.hidden = sourceColumn.hidden;
      }
      if (sourceColumn.outlineLevel !== undefined) {
        targetColumn.outlineLevel = sourceColumn.outlineLevel;
      }
      if (sourceColumn.style !== undefined) {
        targetColumn.style = this.deepClone(sourceColumn.style);
      }
    }

    source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
      const targetRow = target.getRow(rowNumber);
      if (sourceRow.height !== undefined) {
        targetRow.height = sourceRow.height;
      }
      if (sourceRow.hidden !== undefined) {
        targetRow.hidden = sourceRow.hidden;
      }
      if (sourceRow.outlineLevel !== undefined) {
        targetRow.outlineLevel = sourceRow.outlineLevel;
      }

      sourceRow.eachCell({ includeEmpty: true }, (sourceCell, columnNumber) => {
        const targetCell = targetRow.getCell(columnNumber);
        targetCell.value = this.deepClone(sourceCell.value);
        targetCell.style = this.deepClone(sourceCell.style);
        targetCell.dataValidation = this.deepClone(sourceCell.dataValidation);
        if (this.hasMeaningfulNote(sourceCell.note)) {
          targetCell.note = this.deepClone(sourceCell.note);
        }
      });
    });

    const merges = (source as ExcelJsWorksheetWithMerges)._merges ?? {};
    for (const merge of Object.values(merges)) {
      target.mergeCells(merge.range);
    }

    return target;
  }

  private scanWorksheet(worksheet: ExcelJS.Worksheet): WorksheetTemplateSource {
    const rows: RowTemplateSource[] = [];
    const mergeMasters = this.collectMergeMasters(worksheet);

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const cells: CellTemplateSource[] = [];

      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        if (cell.isMerged && cell.master !== cell) {
          return;
        }

        const formula = cell.formula || undefined;
        const sourceRange = mergeMasters.get(`${rowNumber}:${columnNumber}`);
        if (!sourceRange && !formula && this.getTemplateValue(cell) == null && !this.hasStyle(cell.style)) {
          return;
        }

        cells.push({
          address: {
            sheetName: worksheet.name,
            row: rowNumber,
            column: columnNumber,
          },
          ...(sourceRange ? { sourceRange } : {}),
          value: this.getTemplateValue(cell),
          ...(formula ? { formula } : {}),
        });
      });

      if (cells.length > 0 || row.height !== undefined) {
        rows.push({
          rowNumber,
          height: row.height,
          cells,
        });
      }
    });

    return {
      name: worksheet.name,
      sheetId: String(worksheet.id),
      rows,
      merges: [],
    };
  }

  private collectMergeMasters(worksheet: ExcelJS.Worksheet): Map<string, ReturnType<MergeRange['toRangeAddress']>> {
    const masters = new Map<string, ReturnType<MergeRange['toRangeAddress']>>();
    const rawMerges = (worksheet as ExcelJsWorksheetWithMerges)._merges ?? {};
    for (const rawMerge of Object.values(rawMerges)) {
      const range = MergeRange.fromA1(rawMerge.range, worksheet.name).toRangeAddress();
      masters.set(`${range.start.row}:${range.start.column}`, range);
    }

    return masters;
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

  private hasStyle(style: Partial<ExcelJS.Style> | undefined): boolean {
    return !!style && STYLE_KEYS.some((key) => style[key] !== undefined);
  }

  private hasMeaningfulNote(note: ExcelJS.Cell['note']): note is NonNullable<ExcelJS.Cell['note']> {
    if (note === undefined || note === null || note === '') {
      return false;
    }

    if (typeof note === 'string') {
      return note.trim() !== '';
    }

    if (typeof note === 'object' && 'texts' in note && Array.isArray(note.texts)) {
      return note.texts.some((text) => typeof text.text === 'string' && text.text.trim() !== '');
    }

    return true;
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

  private deepClone<T>(value: T): T {
    if (value === undefined || value === null) {
      return value;
    }

    if (typeof structuredClone === 'function') {
      return structuredClone(value);
    }

    return JSON.parse(JSON.stringify(value)) as T;
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

  private sortSheetOperations(operations: readonly RenderOperation[]): readonly RenderOperation[] {
    return [...operations].sort((left, right) => {
      if (left.type === 'CloneColumn' && right.type !== 'CloneColumn') {
        return -1;
      }

      if (right.type === 'CloneColumn' && left.type !== 'CloneColumn') {
        return 1;
      }

      return 0;
    });
  }
}

interface ExcelJsWorksheetWithMerges extends ExcelJS.Worksheet {
  readonly _merges?: Record<string, { readonly range: string }>;
}

const STYLE_KEYS = ['numFmt', 'font', 'fill', 'border', 'alignment', 'protection'] as const;

export interface ExcelJsWorkbookRendererOptions {
  readonly assetResolver?: AssetResolverOptions;
  readonly limits?: Pick<RenderLimits, 'maxTemplateBytes'>;
  readonly preserveFormulas?: boolean;
}
