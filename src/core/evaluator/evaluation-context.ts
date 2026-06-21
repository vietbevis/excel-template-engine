import type { JsonObject } from '../../application/engine/types.js';

export class EvaluationContext {
  private constructor(
    readonly root: JsonObject,
    readonly current: unknown,
    readonly parent?: EvaluationContext,
  ) {}

  static root(data: JsonObject): EvaluationContext {
    return new EvaluationContext(data, data);
  }

  child(current: unknown): EvaluationContext {
    return new EvaluationContext(this.root, current, this);
  }
}
