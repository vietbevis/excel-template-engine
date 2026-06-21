import { JsonPathResolver, type PathResolutionContext } from './json-path-resolver.js';

export interface ExpressionEvaluatorOptions {
  readonly missingValue?: 'empty-string' | 'null' | 'throw';
}

export class ExpressionEvaluator {
  private readonly resolver = new JsonPathResolver();

  evaluate(
    expression: string,
    context: PathResolutionContext,
    options: ExpressionEvaluatorOptions = {},
  ): unknown {
    const parsed = this.parse(expression);
    const value = this.evaluateTerm(parsed.primary, context);

    if (value !== null && value !== undefined) {
      return value;
    }

    if (parsed.fallback) {
      return this.evaluateTerm(parsed.fallback, context);
    }

    if (options.missingValue === 'throw') {
      throw new Error(`Missing value for expression: ${expression}`);
    }

    return options.missingValue === 'null' ? null : '';
  }

  parse(expression: string): ParsedExpression {
    const parts = this.splitNullish(expression.trim());
    const [primary, fallback, ...rest] = parts;

    if (!primary || rest.length > 0) {
      throw new Error(`Invalid expression: ${expression}`);
    }

    const parsedPrimary = this.parseTerm(primary);
    return fallback
      ? { primary: parsedPrimary, fallback: this.parseTerm(fallback) }
      : { primary: parsedPrimary };
  }

  private parseTerm(input: string): ExpressionTerm {
    const value = input.trim();

    if (!value) {
      throw new Error('Expression term cannot be empty.');
    }

    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      return {
        type: 'literal',
        value: this.unquote(value),
      };
    }

    if (value === 'true' || value === 'false') {
      return {
        type: 'literal',
        value: value === 'true',
      };
    }

    if (value === 'null') {
      return {
        type: 'literal',
        value: null,
      };
    }

    const numberValue = Number(value);
    if (/^-?\d+(?:\.\d+)?$/.test(value) && Number.isFinite(numberValue)) {
      return {
        type: 'literal',
        value: numberValue,
      };
    }

    if (!/^[A-Za-z_][A-Za-z0-9_$]*(?:\.[A-Za-z_][A-Za-z0-9_$]*|\[\d+\])*$/.test(value) && value !== '.') {
      throw new Error(`Invalid path expression: ${value}`);
    }

    return {
      type: 'path',
      path: value,
    };
  }

  private evaluateTerm(term: ExpressionTerm, context: PathResolutionContext): unknown {
    if (term.type === 'literal') {
      return term.value;
    }

    return this.resolver.resolve(term.path, context);
  }

  private splitNullish(expression: string): string[] {
    const parts: string[] = [];
    let current = '';
    let quote: '"' | "'" | undefined;

    for (let index = 0; index < expression.length; index += 1) {
      const char = expression[index] || '';
      const next = expression[index + 1] || '';

      if (quote) {
        current += char;
        if (char === quote && expression[index - 1] !== '\\') {
          quote = undefined;
        }
        continue;
      }

      if (char === '"' || char === "'") {
        quote = char;
        current += char;
        continue;
      }

      if (char === '?' && next === '?') {
        parts.push(current.trim());
        current = '';
        index += 1;
        continue;
      }

      current += char;
    }

    if (quote) {
      throw new Error(`Unclosed string literal in expression: ${expression}`);
    }

    parts.push(current.trim());
    return parts;
  }

  private unquote(value: string): string {
    return value
      .slice(1, -1)
      .replace(/\\(["'\\nrt])/g, (_match: string, escaped: string) => {
        if (escaped === 'n') {
          return '\n';
        }
        if (escaped === 'r') {
          return '\r';
        }
        if (escaped === 't') {
          return '\t';
        }
        return escaped;
      });
  }
}

export interface ParsedExpression {
  readonly primary: ExpressionTerm;
  readonly fallback?: ExpressionTerm;
}

export type ExpressionTerm =
  | { readonly type: 'path'; readonly path: string }
  | { readonly type: 'literal'; readonly value: string | number | boolean | null };
