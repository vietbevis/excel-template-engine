import test from 'node:test';
import assert from 'node:assert/strict';
import { LoopContext } from '../src/core/render/loop-context.js';

test('LoopContext expose index, first, last, length và depth', () => {
  const first = LoopContext.forItem(0, 3);
  const middle = LoopContext.forItem(1, 3);
  const last = LoopContext.forItem(2, 3);

  assert.equal(first.index, 0);
  assert.equal(first.first, true);
  assert.equal(first.last, false);
  assert.equal(first.length, 3);
  assert.equal(first.depth, 0);

  assert.equal(middle.first, false);
  assert.equal(middle.last, false);
  assert.equal(last.last, true);
});

test('LoopContext hỗ trợ parent loop resolution', () => {
  const parent = LoopContext.forItem(2, 5);
  const child = LoopContext.forItem(1, 3, parent);

  assert.equal(child.resolve('index'), 1);
  assert.equal(child.resolve('parent.index'), 2);
  assert.equal(child.resolve('loop.last'), false);
  assert.equal(child.depth, 1);
});
