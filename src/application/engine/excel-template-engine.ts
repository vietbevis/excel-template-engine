import { TemplateEngineService } from './template-engine-service.js';
import type {
  EngineHelper,
  EngineRenderOptions,
  EngineRenderRequest,
  EngineRenderResult,
  JsonObject,
  TemplateInput,
} from './types.js';

export class ExcelTemplateEngine {
  private readonly service = new TemplateEngineService();

  registerHelper(name: string, helper: EngineHelper): this {
    this.service.registerHelper(name, helper);
    return this;
  }

  render(
    template: TemplateInput,
    data: JsonObject,
    options: EngineRenderOptions = {},
  ): Promise<EngineRenderResult> {
    return this.service.render({ template, data, options });
  }

  renderRequest(request: EngineRenderRequest): Promise<EngineRenderResult> {
    return this.service.render(request);
  }
}
