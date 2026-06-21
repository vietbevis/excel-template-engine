import type { LoopContext } from '../render/loop-context.js';

export interface PathResolutionContext {
  readonly root: unknown;
  readonly current: unknown;
  readonly loop?: LoopContext | undefined;
}

export class JsonPathResolver {
  resolve(path: string, context: PathResolutionContext): unknown {
    const normalizedPath = path.trim();

    if (!normalizedPath || normalizedPath === '.') {
      return context.current;
    }

    const fromLoop = context.loop?.resolve(normalizedPath);
    if (fromLoop !== undefined) {
      return fromLoop;
    }

    const fromCurrent = this.read(context.current, normalizedPath);
    if (fromCurrent !== undefined) {
      return fromCurrent;
    }

    return this.read(context.root, normalizedPath);
  }

  private read(source: unknown, path: string): unknown {
    const segments = this.toSegments(path);
    let current = source;

    for (const segment of segments) {
      if (current == null) {
        return undefined;
      }

      if (typeof segment === 'number') {
        if (!Array.isArray(current)) {
          return undefined;
        }

        current = current[segment];
        continue;
      }

      if (typeof current !== 'object') {
        return undefined;
      }

      current = (current as Record<string, unknown>)[segment];
    }

    return current;
  }

  private toSegments(path: string): Array<string | number> {
    const segments: Array<string | number> = [];
    let token = '';

    for (let index = 0; index < path.length; index += 1) {
      const char = path[index];

      if (char === '.') {
        if (token) {
          segments.push(token);
          token = '';
        }
        continue;
      }

      if (char === '[') {
        if (token) {
          segments.push(token);
          token = '';
        }

        const end = path.indexOf(']', index);
        if (end === -1) {
          throw new Error(`Invalid JSON path: ${path}`);
        }

        const rawIndex = path.slice(index + 1, end);
        const arrayIndex = Number(rawIndex);
        if (!Number.isInteger(arrayIndex) || arrayIndex < 0) {
          throw new Error(`Invalid array index in JSON path: ${path}`);
        }

        segments.push(arrayIndex);
        index = end;
        continue;
      }

      token += char;
    }

    if (token) {
      segments.push(token);
    }

    return segments;
  }
}
