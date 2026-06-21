import test from 'node:test';
import assert from 'node:assert/strict';
import type {
  BlockNode,
  BlockEndNode,
  BlockStartNode,
  EachColumnNode,
  EachNode,
  HelperNode,
  GridNode,
  IfNode,
  ImageNode,
  PlaceholderNode,
  TextNode,
} from '../src/core/ast/nodes.js';
import { TemplateParser } from '../src/core/parser/template-parser.js';
import { visitTemplateNode, type TemplateNodeVisitor } from '../src/core/visitor/template-node-visitor.js';

test('visitTemplateNode dispatch đúng visitor method cho từng node', () => {
  const parser = new TemplateParser();
  const nodes = parser.parseCell([
    'a',
    '{{name}}',
    '{{sum(scores)}}',
    '{{image avatar}}',
    '{{#if active}}x{{/if}}',
    '{{#each students}}y{{/each}}',
    '{{#each-col subjects}}z{{/each-col}}',
    '{{#block contracts}}w{{/block}}',
    '{{#grid students subjects}}s{{/grid}}',
  ].join(''));

  const visited: string[] = [];
  const visitor: TemplateNodeVisitor<void> = {
    visitText: (_node: TextNode) => { visited.push('text'); },
    visitPlaceholder: (_node: PlaceholderNode) => { visited.push('placeholder'); },
    visitEach: (_node: EachNode) => { visited.push('each'); },
    visitEachColumn: (_node: EachColumnNode) => { visited.push('each-col'); },
    visitGrid: (_node: GridNode) => { visited.push('grid'); },
    visitIf: (_node: IfNode) => { visited.push('if'); },
    visitHelper: (_node: HelperNode) => { visited.push('helper'); },
    visitImage: (_node: ImageNode) => { visited.push('image'); },
    visitBlock: (_node: BlockNode) => { visited.push('block'); },
    visitBlockStart: (_node: BlockStartNode) => { visited.push('block-start'); },
    visitBlockEnd: (_node: BlockEndNode) => { visited.push('block-end'); },
  };

  for (const node of nodes) {
    visitTemplateNode(node, visitor);
  }

  assert.deepEqual(visited, [
    'text',
    'placeholder',
    'helper',
    'image',
    'if',
    'each',
    'each-col',
    'block',
    'grid',
  ]);
});
