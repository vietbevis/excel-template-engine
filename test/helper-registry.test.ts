import test from 'node:test';
import assert from 'node:assert/strict';
import { DefaultHelperRegistry } from '../src/core/evaluator/helper-registry.js';

test('DefaultHelperRegistry gọi helper sync và async', async () => {
  const registry = new DefaultHelperRegistry();

  registry.register('sum', ([values]) => {
    if (!Array.isArray(values)) {
      return 0;
    }

    return values.reduce((total, value) => total + Number(value), 0);
  });

  registry.register('upper', async ([value]) => String(value).toUpperCase());

  const context = {
    data: {},
    root: {},
    current: {},
  };

  assert.equal(await registry.invoke('sum', [[1, 2, 3]], context), 6);
  assert.equal(await registry.invoke('upper', ['abc'], context), 'ABC');
});

test('DefaultHelperRegistry báo lỗi khi helper không tồn tại', async () => {
  const registry = new DefaultHelperRegistry();

  await assert.rejects(
    () => registry.invoke('missing', [], { data: {}, root: {}, current: {} }),
    /Unknown helper/,
  );
});
