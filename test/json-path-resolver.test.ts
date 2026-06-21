import test from 'node:test';
import assert from 'node:assert/strict';
import { EvaluationContext } from '../src/core/evaluator/evaluation-context.js';
import { JsonPathResolver } from '../src/core/evaluator/json-path-resolver.js';

test('JsonPathResolver resolve dot path và bracket index từ root', () => {
  const resolver = new JsonPathResolver();
  const context = EvaluationContext.root({
    teacher: {
      name: 'Nguyễn Văn A',
    },
    students: [
      { name: 'Trần Thị B' },
      { name: 'Lê Văn C' },
    ],
  });

  assert.equal(resolver.resolve('teacher.name', context), 'Nguyễn Văn A');
  assert.equal(resolver.resolve('students[1].name', context), 'Lê Văn C');
});

test('JsonPathResolver ưu tiên scope hiện tại rồi fallback về root', () => {
  const resolver = new JsonPathResolver();
  const root = EvaluationContext.root({
    schoolYear: '2025-2026',
  });
  const child = root.child({
    name: 'Sinh viên A',
  });

  assert.equal(resolver.resolve('name', child), 'Sinh viên A');
  assert.equal(resolver.resolve('schoolYear', child), '2025-2026');
});
