import { DefaultHelperRegistry } from '../evaluator/helper-registry.js';
import { TemplateParser } from '../parser/template-parser.js';
import { RenderContext, type MissingValuePolicy } from '../render/render-context.js';
import { TemplateRenderVisitor } from '../render/template-render-visitor.js';
import type { EngineHelper, JsonObject } from '../../application/engine/types.js';
import type { TemplateNode } from '../ast/nodes.js';

export interface CoreTemplateEngineOptions {
  readonly missingValue?: MissingValuePolicy;
}

export class TemplateEngine {
  private readonly parser = new TemplateParser();
  private readonly helpers = new DefaultHelperRegistry();

  registerHelper(name: string, helper: EngineHelper): this {
    this.helpers.register(name, helper);
    return this;
  }

  parse(template: string): readonly TemplateNode[] {
    return this.parser.parseCell(template);
  }

  async render(
    template: string,
    data: JsonObject,
    options: CoreTemplateEngineOptions = {},
  ): Promise<string> {
    const ast = this.parse(template);
    return this.renderNodes(ast, data, options);
  }

  async renderNodes(
    nodes: readonly TemplateNode[],
    data: JsonObject,
    options: CoreTemplateEngineOptions = {},
  ): Promise<string> {
    const contextOptions = options.missingValue
      ? { helpers: this.helpers, missingValue: options.missingValue }
      : { helpers: this.helpers };
    const context = RenderContext.create(data, {
      ...contextOptions,
    });
    return new TemplateRenderVisitor(context).render(nodes);
  }
}
