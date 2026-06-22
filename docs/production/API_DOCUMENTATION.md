# API Documentation

Date: 2026-06-21

## Install

```bash
npm install excel-template-engine
```

## API Stability

Version `0.1.x` treats `ExcelTemplateEngine`, `EngineRenderOptions`, `EngineRenderResult`, helper registration, render limits, and asset resolver options as the stable public API. AST nodes, render-plan operations, and ExcelJS infrastructure classes are exported for advanced users and tests, but may change before `1.0.0`.

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
  readonly limits?: RenderLimits;
  readonly assetResolver?: AssetResolverOptions;
}
```

Current note: `preserveFormulas` and `recalculateFormulas` are declared but not fully enforced as production policies yet. Formula preservation/shift behavior is implemented in ExcelJS clone paths.

Render limits:

```ts
interface RenderLimits {
  readonly maxTemplateBytes?: number;
  readonly maxWorksheets?: number;
  readonly maxRows?: number;
  readonly maxColumns?: number;
  readonly maxOperations?: number;
}
```

Asset resolver options:

```ts
interface AssetResolverOptions {
  readonly baseDir?: string;
  readonly allowAbsolutePaths?: boolean;
  readonly maxBytes?: number;
}
```

Defaults:

- image paths are resolved relative to the template file directory;
- absolute image paths are disabled by default;
- image sources are limited to 10 MiB by default.
- `maxTemplateBytes` rejects oversized template files or buffers before ExcelJS loads them.

Example:

```ts
await engine.render('template.xlsx', data, {
  limits: {
    maxTemplateBytes: 5 * 1024 * 1024,
    maxWorksheets: 50,
    maxRows: 100_000,
    maxColumns: 5_000,
    maxOperations: 200_000,
  },
  assetResolver: {
    maxBytes: 2 * 1024 * 1024,
  },
});
```

Limit violations throw `LimitExceededError`:

```ts
{
  code: 'LIMIT_EXCEEDED',
  details: {
    limitName: string,
    actual: number,
    limit: number
  }
}
```

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

Dynamic grouped column header:

```text
{{#each-col groups span=size}}{{name}}{{/each-col}}
{{#each-col groups span=size rowspan=2}}{{name}}{{/each-col}}
{{#each-col groups span=size rowspan=2 reserve=31}}{{name}}{{/each-col}}
```

`span` reads an integer field from each item and merges the rendered cell horizontally by that number of columns. Missing/null `span` falls back to `1`; `0` or negative spans skip the item. `rowspan` is an integer literal for vertical merge height and defaults to `1`. `reserve` is an integer literal for templates that already contain a reserved dynamic column area; the renderer fills the reserved columns and only clones more columns if the rendered width exceeds the reserve.

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
