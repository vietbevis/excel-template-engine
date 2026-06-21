import { DefaultHelperRegistry } from '../evaluator/helper-registry.js';
import type { EngineHelper, JsonObject } from '../../application/engine/types.js';
import type { LoopContext } from './loop-context.js';

export type MissingValuePolicy = 'empty-string' | 'null' | 'throw';

export interface RenderContextOptions {
  readonly missingValue?: MissingValuePolicy;
  readonly helpers?: DefaultHelperRegistry;
}

export class RenderContext {
  readonly helpers: DefaultHelperRegistry;
  readonly missingValue: MissingValuePolicy;

  private constructor(
    readonly root: JsonObject,
    readonly current: unknown,
    options: RenderContextOptions,
    readonly parent?: RenderContext,
    readonly loop?: LoopContext,
  ) {
    this.helpers = options.helpers ?? new DefaultHelperRegistry();
    this.missingValue = options.missingValue ?? 'empty-string';
  }

  static create(data: JsonObject, options: RenderContextOptions = {}): RenderContext {
    return new RenderContext(data, data, options);
  }

  child(current: unknown, loop?: LoopContext): RenderContext {
    return new RenderContext(
      this.root,
      current,
      {
        helpers: this.helpers,
        missingValue: this.missingValue,
      },
      this,
      loop,
    );
  }

  registerHelper(name: string, helper: EngineHelper): void {
    this.helpers.register(name, helper);
  }
}
