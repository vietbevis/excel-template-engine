import { DefaultHelperRegistry } from '../../core/evaluator/helper-registry.js';
import { EvaluationContext } from '../../core/evaluator/evaluation-context.js';
import { TemplateParser } from '../../core/parser/template-parser.js';
import { RenderPlanner } from '../planner/render-planner.js';
import { ExcelJsWorkbookRenderer } from '../../infrastructure/exceljs/excel-js-workbook-renderer.js';
import type { EngineHelper, EngineRenderRequest, EngineRenderResult } from './types.js';

export class TemplateEngineService {
  private readonly helpers = new DefaultHelperRegistry();
  private readonly parser = new TemplateParser();
  private readonly planner = new RenderPlanner(this.helpers);

  registerHelper(name: string, helper: EngineHelper): void {
    this.helpers.register(name, helper);
  }

  async render(request: EngineRenderRequest): Promise<EngineRenderResult> {
    const renderer = request.options?.renderer ?? new ExcelJsWorkbookRenderer(
      {
        ...(request.options?.assetResolver ? { assetResolver: request.options.assetResolver } : {}),
        ...(request.options?.limits?.maxTemplateBytes !== undefined
          ? { limits: { maxTemplateBytes: request.options.limits.maxTemplateBytes } }
          : {}),
      },
    );
    const workbookSource = await renderer.load(request.template);
    const ast = this.parser.parseWorkbook(workbookSource);
    const context = EvaluationContext.root(request.data);
    const plan = await this.planner.createPlan(ast, context, request.options);
    await renderer.apply(plan);

    return {
      output: await renderer.write(),
      warnings: plan.warnings,
    };
  }
}
