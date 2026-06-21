# API Documentation

Date: 2026-06-21

## Install

```bash
npm install excel-template-engine
```

## Basic Usage

```ts
import { writeFile } from 'node:fs/promises';
import { ExcelTemplateEngine } from 'excel-template-engine';

const engine = new ExcelTemplateEngine();

const result = await engine.render('template.xlsx', {
  teacher: { name: 'Nguyen Van A' },
  contract: { code: 'HD-2026' },
});

await writeFile('output.xlsx', result.output);
```

## Public Classes

### `ExcelTemplateEngine`

Main XLSX rendering API.

```ts
const engine = new ExcelTemplateEngine();
```

Methods:

- `registerHelper(name, helper): this`
- `render(template, data, options?): Promise<EngineRenderResult>`
- `renderRequest(request): Promise<EngineRenderResult>`

### `TemplateEngine`

Core string-template renderer useful for testing parser/evaluator behavior without ExcelJS.

```ts
const engine = new TemplateEngine();
const output = await engine.render('Hello {{user.name}}', { user: { name: 'A' } });
```

## Template Input

`template` can be:

- file path string;
- `Buffer`;
- `Uint8Array`;
- `ArrayBuffer`.

## Render Result

```ts
interface EngineRenderResult {
  readonly output: Uint8Array;
  readonly warnings: readonly string[];
}
```

## Options

```ts
interface EngineRenderOptions {
  readonly preserveFormulas?: boolean;
  readonly recalculateFormulas?: boolean;
  readonly missingValue?: 'empty-string' | 'null' | 'throw';
  readonly renderer?: WorkbookRenderer;
}
```

Current note: `preserveFormulas` and `recalculateFormulas` are declared but not fully enforced as production policies yet. Formula preservation/shift behavior is implemented in ExcelJS clone paths.

## Helpers

```ts
engine.registerHelper('currency', ([value]) => `${Number(value).toLocaleString('en-US')} VND`);
engine.registerHelper('sum', ([values]) => Array.isArray(values)
  ? values.reduce((total, value) => total + Number(value), 0)
  : 0);
```

Template:

```text
{{currency(price)}}
{{sum(scores)}}
```

Helper signature:

```ts
type EngineHelper = (
  args: readonly unknown[],
  context: HelperContext,
) => unknown | Promise<unknown>;
```

## Supported Syntax

Placeholder:

```text
{{teacher.name}}
{{teacher.name ?? "Unknown"}}
```

Row loop:

```text
{{#each students}}{{name}}{{/each}}
```

Loop context:

```text
{{index}}
{{first}}
{{last}}
{{parent.index}}
```

Column loop:

```text
{{#each-col subjects}}{{name}}{{/each-col}}
```

Grid:

```text
{{#grid students subjects}}{{score}}{{/grid}}
```

Block:

```text
{{#block contracts}}
{{code}}
{{amount}}
{{/block}}
```

Image:

```text
{{image avatar}}
```

Image sources:

- png path;
- jpg/jpeg path;
- `Buffer`;
- `Uint8Array`;
- `ArrayBuffer`;
- base64 string;
- image data URL.

## Missing Values

```ts
await engine.render(template, data, {
  missingValue: 'throw',
});
```

Policies:

- `empty-string`: default empty output;
- `null`: evaluates missing values as null, stringified to empty;
- `throw`: throws an error.

## Typedoc

Generate API reference:

```bash
npm run docs:api
```

Output directory:

```text
docs/api
```
