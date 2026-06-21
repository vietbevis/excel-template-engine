import test from 'node:test';
import assert from 'node:assert/strict';
import { TemplateEngine } from '../src/core/engine/template-engine.js';

test('TemplateEngine render placeholder, helper và if', async () => {
  const engine = new TemplateEngine();
  engine.registerHelper('sum', ([values]) => {
    if (!Array.isArray(values)) {
      return 0;
    }

    return values.reduce((total, value) => total + Number(value), 0);
  });

  const output = await engine.render(
    'GV: {{teacher.name}} | Tổng: {{sum(scores)}} | {{#if hasSignature}}Đã ký{{/if}}',
    {
      teacher: { name: 'Nguyễn Văn A' },
      scores: [8, 9, 10],
      hasSignature: true,
    },
  );

  assert.equal(output, 'GV: Nguyễn Văn A | Tổng: 27 | Đã ký');
});

test('TemplateEngine render placeholder nested path và default value', async () => {
  const engine = new TemplateEngine();

  assert.equal(
    await engine.render('{{teacher.name}} {{contract.code}} {{user.profile.email}}', {
      teacher: { name: 'Nguyễn Văn A' },
      contract: { code: 'HD001' },
      user: { profile: { email: 'user@example.com' } },
    }),
    'Nguyễn Văn A HD001 user@example.com',
  );

  assert.equal(
    await engine.render('{{teacher.name ?? "Unknown"}}', {
      teacher: null,
    }),
    'Unknown',
  );
});

test('TemplateEngine render each, each-col và block theo scope hiện tại', async () => {
  const engine = new TemplateEngine();

  const output = await engine.render(
    [
      '{{#each students}}{{code}}:{{name}};{{/each}}',
      '{{#each-col subjects}}[{{name}}]{{/each-col}}',
      '{{#block contracts}}{{number}}|{{/block}}',
    ].join(''),
    {
      students: [
        { code: 'SV001', name: 'Trần Thị B' },
        { code: 'SV002', name: 'Lê Văn C' },
      ],
      subjects: [{ name: 'Toán' }, { name: 'Lý' }],
      contracts: [{ number: 'HD01' }, { number: 'HD02' }],
    },
  );

  assert.equal(output, 'SV001:Trần Thị B;SV002:Lê Văn C;[Toán][Lý]HD01|HD02|');
});

test('TemplateEngine render EachNode với index, first, last', async () => {
  const engine = new TemplateEngine();

  const output = await engine.render(
    '{{#each students}}{{index}}:{{name}}:{{first}}:{{last}};{{/each}}',
    {
      students: [
        { name: 'A' },
        { name: 'B' },
        { name: 'C' },
      ],
    },
  );

  assert.equal(output, '0:A:true:false;1:B:false:false;2:C:false:true;');
});

test('TemplateEngine render nested each với index riêng cho từng loop', async () => {
  const engine = new TemplateEngine();

  const output = await engine.render(
    [
      '{{#each classes}}',
      'class{{index}}[',
      '{{#each students}}{{parent.index}}.{{index}}:{{name}}:{{first}}:{{last}};',
      '{{/each}}',
      ']',
      '{{/each}}',
    ].join(''),
    {
      classes: [
        {
          students: [
            { name: 'A' },
            { name: 'B' },
          ],
        },
        {
          students: [
            { name: 'C' },
          ],
        },
      ],
    },
  );

  assert.equal(output, 'class0[0.0:A:true:false;0.1:B:false:true;]class1[1.0:C:true:true;]');
});

test('TemplateEngine render image node bằng giá trị path ở core layer', async () => {
  const engine = new TemplateEngine();

  const output = await engine.render('{{image avatar}}', {
    avatar: '/tmp/avatar.png',
  });

  assert.equal(output, '/tmp/avatar.png');
});

test('TemplateEngine hỗ trợ missing value policy', async () => {
  const engine = new TemplateEngine();

  assert.equal(await engine.render('{{missing}}', {}, { missingValue: 'empty-string' }), '');
  assert.equal(await engine.render('{{missing}}', {}, { missingValue: 'null' }), '');
  await assert.rejects(
    () => engine.render('{{missing}}', {}, { missingValue: 'throw' }),
    /Missing value/,
  );
});
