import { performance } from 'node:perf_hooks';
import { ExpressionEvaluator } from '../src/core/evaluator/expression-evaluator.js';
import { TemplateEngine } from '../src/core/engine/template-engine.js';

const iterations = 50_000;
const data = {
  teacher: {
    name: 'Nguyễn Văn A',
  },
  contract: {
    code: 'HD001',
  },
  user: {
    profile: {
      email: 'user@example.com',
    },
  },
};

const evaluator = new ExpressionEvaluator();
const evaluatorStart = performance.now();

for (let index = 0; index < iterations; index += 1) {
  evaluator.evaluate('user.profile.email ?? "Unknown"', {
    current: data,
    root: data,
  });
}

const evaluatorDuration = performance.now() - evaluatorStart;

const engine = new TemplateEngine();
const template = '{{teacher.name}} {{contract.code}} {{user.profile.email ?? "Unknown"}}';
const renderStart = performance.now();

for (let index = 0; index < iterations; index += 1) {
  await engine.render(template, data);
}

const renderDuration = performance.now() - renderStart;

console.log(JSON.stringify({
  iterations,
  evaluator: {
    durationMs: Math.round(evaluatorDuration),
    opsPerSecond: Math.round(iterations / (evaluatorDuration / 1000)),
  },
  placeholderRender: {
    durationMs: Math.round(renderDuration),
    opsPerSecond: Math.round(iterations / (renderDuration / 1000)),
  },
}, null, 2));
