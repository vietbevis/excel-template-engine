import { DefaultHelperRegistry } from '../../core/evaluator/helper-registry.js';
import { EvaluationContext } from '../../core/evaluator/evaluation-context.js';
import { TemplateParser } from '../../core/parser/template-parser.js';
import { RenderPlanner } from '../planner/render-planner.js';
import { ExcelJsWorkbookRenderer } from '../../infrastructure/exceljs/excel-js-workbook-renderer.js';
import { JsonPathResolver } from '../../core/evaluator/json-path-resolver.js';
import type {
  EngineHelper,
  EngineRenderRequest,
  EngineRenderResult,
  JsonObject,
  WorkbookRenderConfig,
} from './types.js';

export class TemplateEngineService {
  private readonly helpers = new DefaultHelperRegistry();
  private readonly parser = new TemplateParser();
  private readonly planner = new RenderPlanner(this.helpers);
  private readonly pathResolver = new JsonPathResolver();

  registerHelper(name: string, helper: EngineHelper): void {
    this.helpers.register(name, helper);
  }

  async render(request: EngineRenderRequest): Promise<EngineRenderResult> {
    const renderer = request.options?.renderer ?? new ExcelJsWorkbookRenderer(
      {
        ...(request.options?.preserveFormulas !== undefined ? { preserveFormulas: request.options.preserveFormulas } : {}),
        ...(request.options?.assetResolver ? { assetResolver: request.options.assetResolver } : {}),
        ...(request.options?.limits?.maxTemplateBytes !== undefined
          ? { limits: { maxTemplateBytes: request.options.limits.maxTemplateBytes } }
          : {}),
      },
    );
    let workbookSource = await renderer.load(request.template);
    const workbookConfig = this.resolveWorkbookConfig(request);
    if (workbookConfig && renderer.prepare) {
      workbookSource = await renderer.prepare(workbookConfig);
    }
    const ast = this.parser.parseWorkbook(workbookSource);
    const context = EvaluationContext.root(request.data);
    const plan = await this.planner.createPlan(
      ast,
      context,
      request.options,
      this.createSheetContexts(request.data, workbookConfig),
    );
    await renderer.apply(plan);

    return {
      output: await renderer.write(),
      warnings: plan.warnings,
    };
  }

  private resolveWorkbookConfig(request: EngineRenderRequest): WorkbookRenderConfig | undefined {
    if (request.options?.workbook) {
      return request.options.workbook;
    }

    const rawConfig = request.data._workbook;
    if (this.isWorkbookRenderConfig(rawConfig)) {
      return rawConfig;
    }

    return undefined;
  }

  private createSheetContexts(
    data: JsonObject,
    config: WorkbookRenderConfig | undefined,
  ): ReadonlyMap<string, EvaluationContext> {
    const contexts = new Map<string, EvaluationContext>();
    for (const worksheet of config?.worksheets ?? []) {
      if (!worksheet.dataPath) {
        continue;
      }

      const scopedData = this.pathResolver.resolve(worksheet.dataPath, {
        root: data,
        current: data,
      });
      contexts.set(
        worksheet.name,
        EvaluationContext.root(this.toJsonObject(scopedData, worksheet.dataPath)),
      );
    }

    return contexts;
  }

  private toJsonObject(value: unknown, path: string): JsonObject {
    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      return value as JsonObject;
    }

    throw new Error(`Worksheet dataPath must resolve to an object: ${path}`);
  }

  private isWorkbookRenderConfig(value: unknown): value is WorkbookRenderConfig {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) {
      return false;
    }

    const worksheets = (value as WorkbookRenderConfig).worksheets;
    return worksheets === undefined || Array.isArray(worksheets);
  }
}
