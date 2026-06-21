# Migration Guide

Date: 2026-06-21

## From `xlsx-template`

`excel-template-engine` is not a drop-in replacement for `xlsx-template`. It uses a TypeScript-first API, a restricted expression language, explicit helper registration, and an internal render-plan pipeline.

## Basic Rendering

Before:

```js
// xlsx-template style APIs vary by usage.
```

After:

```ts
import { ExcelTemplateEngine } from 'excel-template-engine';

const engine = new ExcelTemplateEngine();
const result = await engine.render('template.xlsx', data);
```

## Placeholder Syntax

Use:

```text
{{teacher.name}}
{{contract.code}}
{{user.profile.email}}
```

Default values:

```text
{{teacher.name ?? "Unknown"}}
```

## Helpers

Register helpers in code:

```ts
engine.registerHelper('sum', ([values]) => Array.isArray(values)
  ? values.reduce((total, value) => total + Number(value), 0)
  : 0);
```

Template:

```text
{{sum(scores)}}
```

## Loops

Rows:

```text
{{#each students}}
{{name}}
{{/each}}
```

Columns:

```text
{{#each-col subjects}}{{name}}{{/each-col}}
```

Blocks:

```text
{{#block contracts}}
{{code}}
{{amount}}
{{/block}}
```

## Images

Template:

```text
{{image avatar}}
```

Data:

```json
{
  "avatar": "./avatar.png"
}
```

Supported sources:

- png;
- jpg/jpeg;
- Buffer;
- Uint8Array;
- ArrayBuffer;
- base64;
- data URL.

## Formulas

The engine can preserve formulas and shift A1 references during clone operations. Supported reference forms:

- `A1`
- `$A$1`
- `A$1`
- `$A1`

Cross-sheet references are supported for simple A1 references such as:

```text
Summary!$B$1+B3
```

## Breaking Differences

- Template expressions are not JavaScript. `eval` and `Function` are intentionally unsupported.
- Missing values default to empty string unless `missingValue: 'throw'` is configured.
- Templates should be treated as trusted input until formula and asset policies are hardened.
- Large-file support must be verified against explicit limits before migration.

## Migration Checklist

- [ ] Convert placeholders to `{{path.to.value}}`.
- [ ] Move custom logic into registered helpers.
- [ ] Replace large dynamic regions with `each`, `each-col`, `grid`, or `block`.
- [ ] Verify merged cells and formulas after rendering.
- [ ] Add regression tests for every template family.
- [ ] Run `npm test` and project-specific smoke tests.
