export interface LoopContextState {
  readonly index: number;
  readonly length: number;
  readonly parent?: LoopContext;
}

export class LoopContext {
  readonly index: number;
  readonly length: number;
  readonly first: boolean;
  readonly last: boolean;
  readonly depth: number;
  readonly parent: LoopContext | undefined;

  constructor(state: LoopContextState) {
    this.index = state.index;
    this.length = state.length;
    this.first = state.index === 0;
    this.last = state.index === state.length - 1;
    this.depth = state.parent ? state.parent.depth + 1 : 0;
    this.parent = state.parent;
  }

  static forItem(index: number, length: number, parent?: LoopContext): LoopContext {
    return new LoopContext(parent ? { index, length, parent } : { index, length });
  }

  resolve(path: string): unknown {
    if (path === 'index') {
      return this.index;
    }

    if (path === 'first') {
      return this.first;
    }

    if (path === 'last') {
      return this.last;
    }

    if (path === 'length') {
      return this.length;
    }

    if (path === 'depth') {
      return this.depth;
    }

    if (path.startsWith('loop.')) {
      return this.resolve(path.slice('loop.'.length));
    }

    if (path.startsWith('parent.') && this.parent) {
      return this.parent.resolve(path.slice('parent.'.length));
    }

    return undefined;
  }
}
