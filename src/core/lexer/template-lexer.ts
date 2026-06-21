import type { Token } from './tokens.js';

export class TemplateLexer {
  tokenize(input: string): readonly Token[] {
    const tokens: Token[] = [];
    let cursor = 0;

    while (cursor < input.length) {
      const open = input.indexOf('{{', cursor);

      if (open === -1) {
        this.pushText(tokens, input.slice(cursor), cursor);
        cursor = input.length;
        break;
      }

      if (open > cursor) {
        this.pushText(tokens, input.slice(cursor, open), cursor);
      }

      const close = input.indexOf('}}', open + 2);
      if (close === -1) {
        throw new Error('Unclosed template expression.');
      }

      tokens.push({ type: 'OpenExpression', value: '{{', span: { start: open, end: open + 2 } });
      this.tokenizeExpression(input.slice(open + 2, close), open + 2, tokens);
      tokens.push({ type: 'CloseExpression', value: '}}', span: { start: close, end: close + 2 } });
      cursor = close + 2;
    }

    tokens.push({ type: 'EOF', value: '', span: { start: input.length, end: input.length } });
    return tokens;
  }

  private pushText(tokens: Token[], value: string, start: number): void {
    if (!value) {
      return;
    }

    tokens.push({
      type: 'Text',
      value,
      span: {
        start,
        end: start + value.length,
      },
    });
  }

  private tokenizeExpression(expression: string, offset: number, tokens: Token[]): void {
    let cursor = 0;

    while (cursor < expression.length) {
      const char = expression[cursor] || '';
      const absolute = offset + cursor;

      if (/\s/.test(char)) {
        cursor += 1;
        continue;
      }

      if (char === '#') {
        tokens.push({ type: 'Hash', value: char, span: { start: absolute, end: absolute + 1 } });
        cursor += 1;
        continue;
      }

      if (char === '/') {
        tokens.push({ type: 'Slash', value: char, span: { start: absolute, end: absolute + 1 } });
        cursor += 1;
        continue;
      }

      if (char === '(') {
        tokens.push({ type: 'OpenParen', value: char, span: { start: absolute, end: absolute + 1 } });
        cursor += 1;
        continue;
      }

      if (char === ')') {
        tokens.push({ type: 'CloseParen', value: char, span: { start: absolute, end: absolute + 1 } });
        cursor += 1;
        continue;
      }

      if (char === ',') {
        tokens.push({ type: 'Comma', value: char, span: { start: absolute, end: absolute + 1 } });
        cursor += 1;
        continue;
      }

      if (char === '"' || char === "'") {
        const quote = char;
        let end = cursor + 1;
        while (end < expression.length && expression[end] !== quote) {
          end += 1;
        }
        if (end >= expression.length) {
          throw new Error('Unclosed string literal.');
        }
        tokens.push({
          type: 'String',
          value: expression.slice(cursor + 1, end),
          span: { start: absolute, end: offset + end + 1 },
        });
        cursor = end + 1;
        continue;
      }

      const numberMatch = /^\d+(?:\.\d+)?/.exec(expression.slice(cursor));
      if (numberMatch) {
        tokens.push({
          type: 'Number',
          value: numberMatch[0],
          span: { start: absolute, end: absolute + numberMatch[0].length },
        });
        cursor += numberMatch[0].length;
        continue;
      }

      const wordMatch = /^[A-Za-z_][A-Za-z0-9_.\-[\]]*/.exec(expression.slice(cursor));
      if (wordMatch) {
        const value = wordMatch[0];
        const type = value.includes('.') || value.includes('[') ? 'Path' : this.wordTokenType(value);
        tokens.push({
          type,
          value,
          span: { start: absolute, end: absolute + value.length },
        });
        cursor += value.length;
        continue;
      }

      throw new Error(`Unexpected token in template expression: ${char}`);
    }
  }

  private wordTokenType(value: string): Token['type'] {
    if (value === 'true' || value === 'false') {
      return 'Boolean';
    }

    if (value === 'null') {
      return 'Null';
    }

    return 'Identifier';
  }
}
