import test from 'node:test';
import assert from 'node:assert/strict';
import { TemplateLexer } from '../src/core/lexer/template-lexer.js';

test('TemplateLexer tách text và expression token', () => {
  const lexer = new TemplateLexer();
  const tokens = lexer.tokenize('Xin chào {{teacher.name}}');

  assert.deepEqual(
    tokens.map((token) => [token.type, token.value]),
    [
      ['Text', 'Xin chào '],
      ['OpenExpression', '{{'],
      ['Path', 'teacher.name'],
      ['CloseExpression', '}}'],
      ['EOF', ''],
    ],
  );
});

test('TemplateLexer nhận diện control tag và helper token', () => {
  const lexer = new TemplateLexer();
  const tokens = lexer.tokenize('{{#each students}}{{sum(scores)}}{{/each}}');

  assert.deepEqual(
    tokens.map((token) => token.type),
    [
      'OpenExpression',
      'Hash',
      'Identifier',
      'Identifier',
      'CloseExpression',
      'OpenExpression',
      'Identifier',
      'OpenParen',
      'Identifier',
      'CloseParen',
      'CloseExpression',
      'OpenExpression',
      'Slash',
      'Identifier',
      'CloseExpression',
      'EOF',
    ],
  );
});
