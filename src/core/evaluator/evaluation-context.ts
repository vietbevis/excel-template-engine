import type { JsonObject } from '../../application/engine/types.js';
import type { LoopContext } from '../render/loop-context.js';

export class EvaluationContext {
  private constructor(
    readonly root: JsonObject,
    readonly current: unknown,
    readonly parent?: EvaluationContext,
    readonly loop?: LoopContext,
  ) {}

  static root(data: JsonObject): EvaluationContext {
    return new EvaluationContext(data, data);
  }

  child(current: unknown, loop: LoopContext | undefined = this.loop): EvaluationContext {
    return new EvaluationContext(this.root, current, this, loop);
  }
}
