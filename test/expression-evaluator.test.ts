import test from 'node:test';
import assert from 'node:assert/strict';
import { ExpressionEvaluator } from '../src/core/evaluator/expression-evaluator.js';

const evaluator = new ExpressionEvaluator();

test('ExpressionEvaluator resolve nested object path', () => {
  const value = evaluator.evaluate('user.profile.email', {
    current: {},
    root: {
      user: {
        profile: {
          email: 'user@example.com',
        },
      },
    },
  });

  assert.equal(value, 'user@example.com');
});

test('ExpressionEvaluator null safe khi path không tồn tại', () => {
  const value = evaluator.evaluate('teacher.profile.name', {
    current: {},
    root: {
      teacher: null,
    },
  });

  assert.equal(value, '');
});

test('ExpressionEvaluator hỗ trợ default value bằng nullish coalescing', () => {
  const value = evaluator.evaluate('teacher.name ?? "Unknown"', {
    current: {},
    root: {
      teacher: {},
    },
  });

  assert.equal(value, 'Unknown');
});

test('ExpressionEvaluator chỉ dùng fallback khi giá trị null hoặc undefined', () => {
  const context = {
    current: {},
    root: {
      teacher: {
        name: '',
        active: false,
        score: 0,
      },
    },
  };

  assert.equal(evaluator.evaluate('teacher.name ?? "Unknown"', context), '');
  assert.equal(evaluator.evaluate('teacher.active ?? true', context), false);
  assert.equal(evaluator.evaluate('teacher.score ?? 10', context), 0);
});

test('ExpressionEvaluator không cho expression không hợp lệ', () => {
  assert.throws(
    () => evaluator.evaluate('teacher.name + other', { current: {}, root: {} }),
    /Invalid path expression/,
  );
});
