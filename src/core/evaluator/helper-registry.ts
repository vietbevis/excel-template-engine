import type { EngineHelper, HelperContext } from '../../application/engine/types.js';

export class DefaultHelperRegistry {
  private readonly helpers = new Map<string, EngineHelper>();

  register(name: string, helper: EngineHelper): void {
    this.helpers.set(name, helper);
  }

  async invoke(name: string, args: readonly unknown[], context: HelperContext): Promise<unknown> {
    const helper = this.helpers.get(name);
    if (!helper) {
      throw new Error(`Unknown helper: ${name}`);
    }

    return helper(args, context);
  }

  has(name: string): boolean {
    return this.helpers.has(name);
  }
}
