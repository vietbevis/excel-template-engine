import test from 'node:test';
import assert from 'node:assert/strict';
import { TemplateParser } from '../src/core/parser/template-parser.js';

test('TemplateParser parse text và placeholder node', () => {
  const parser = new TemplateParser();
  const nodes = parser.parseCell('Giảng viên: {{teacher.name}}');

  assert.equal(nodes.length, 2);
  assert.equal(nodes[0]?.kind, 'TextNode');
  assert.equal(nodes[1]?.kind, 'PlaceholderNode');
  assert.equal(nodes[1]?.kind === 'PlaceholderNode' ? nodes[1].path : undefined, 'teacher.name');
});

test('TemplateParser parse helper và image node', () => {
  const parser = new TemplateParser();
  const nodes = parser.parseCell('{{sum(scores)}} {{image avatar}}');

  assert.equal(nodes[0]?.kind, 'HelperNode');
  assert.equal(nodes[0]?.kind === 'HelperNode' ? nodes[0].name : undefined, 'sum');
  assert.equal(nodes[2]?.kind, 'ImageNode');
  assert.equal(nodes[2]?.kind === 'ImageNode' ? nodes[2].path : undefined, 'avatar');
});

test('TemplateParser parse each lồng if', () => {
  const parser = new TemplateParser();
  const nodes = parser.parseCell('{{#each students}}{{#if active}}{{name}}{{/if}}{{/each}}');

  assert.equal(nodes.length, 1);
  assert.equal(nodes[0]?.kind, 'EachNode');

  if (nodes[0]?.kind !== 'EachNode') {
    assert.fail('Expected EachNode');
  }

  assert.equal(nodes[0].path, 'students');
  assert.equal(nodes[0].children[0]?.kind, 'IfNode');
});

test('TemplateParser parse each-col, block và image node', () => {
  const parser = new TemplateParser();
  const columnNodes = parser.parseCell('{{#each-col subjects}}{{name}}{{/each-col}}');
  const blockNodes = parser.parseCell('{{#block contracts}}{{code}}{{/block}}');
  const imageNodes = parser.parseCell('{{image avatar}}');

  assert.equal(columnNodes[0]?.kind, 'EachColumnNode');
  assert.equal(columnNodes[0]?.kind === 'EachColumnNode' ? columnNodes[0].path : undefined, 'subjects');
  assert.equal(blockNodes[0]?.kind, 'BlockNode');
  assert.equal(blockNodes[0]?.kind === 'BlockNode' ? blockNodes[0].path : undefined, 'contracts');
  assert.equal(imageNodes[0]?.kind, 'ImageNode');
  assert.equal(imageNodes[0]?.kind === 'ImageNode' ? imageNodes[0].path : undefined, 'avatar');
});

test('TemplateParser parse grid node với rowPath và columnPath', () => {
  const parser = new TemplateParser();
  const nodes = parser.parseCell('{{#grid students subjects}}{{score}}{{/grid}}');

  assert.equal(nodes[0]?.kind, 'GridNode');
  assert.equal(nodes[0]?.kind === 'GridNode' ? nodes[0].rowPath : undefined, 'students');
  assert.equal(nodes[0]?.kind === 'GridNode' ? nodes[0].columnPath : undefined, 'subjects');
});

test('TemplateParser báo lỗi khi thiếu closing tag', () => {
  const parser = new TemplateParser();

  assert.throws(
    () => parser.parseCell('{{#each students}}{{name}}'),
    /Missing closing tag/,
  );
});
