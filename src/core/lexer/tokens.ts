import type { SourceSpan } from '../ast/nodes.js';

export type TokenType =
  | 'Text'
  | 'OpenExpression'
  | 'CloseExpression'
  | 'Identifier'
  | 'Path'
  | 'String'
  | 'Number'
  | 'Boolean'
  | 'Null'
  | 'OpenParen'
  | 'CloseParen'
  | 'Comma'
  | 'Hash'
  | 'Slash'
  | 'EOF';

export interface Token {
  readonly type: TokenType;
  readonly value: string;
  readonly span: SourceSpan;
}
