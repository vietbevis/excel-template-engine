# Architecture Review

Date: 2026-06-21

## Executive Summary

The project is moving in the right direction for a production TypeScript XLSX template engine: core parsing/evaluation is separated from ExcelJS infrastructure, render intent is expressed through `RenderPlan`, and the public API is small. Current tests cover placeholders, helpers, loops, blocks, images, formula shifting, merge/style cloning, and multi-sheet rendering.

The main production gaps are no longer feature discovery; they are API stabilization, stronger package boundaries, and bounded-memory guarantees for complex large workbooks. The current ExcelJS backend has benchmark evidence for minimal 100k-row, 5k-column, and 50-worksheet renders, but production policy should still define measured limits for styled, merged, image-heavy, and formula-heavy templates.

## Evidence Reviewed

- Source files: 40 TypeScript source files under `src`.
- Tests: 22 TypeScript tests under `test`.
- Verification run: `npm run typecheck` passed.
- Verification run: `npm test` passed, 56/56 tests.
- Dependency audit: `npm audit --omit=dev` reported 0 vulnerabilities.
- Package dry run after `build` and Typedoc generation: `npm pack --dry-run` produced a 172.9 kB package with 293 files.
- Dependency graph scan: 40 files, 98 internal edges, 0 cycles.

## Layering

Current layers:

- `src/core`: AST, parser, evaluator, render visitor, merge and grid domain helpers.
- `src/application`: engine service, ports, render planner, render plan operations.
- `src/infrastructure`: ExcelJS workbook/worksheet/render managers and default asset resolver.
- `src/shared`: address parsing, errors, common result types.

Strengths:

- The parser/evaluator do not depend on ExcelJS.
- ExcelJS details are concentrated in infrastructure classes.
- `WorkbookRenderer`, `StyleCloneManager`, `MergeManager`, `FormulaManager`, `ImageManager`, and `AssetResolver` ports make the design extensible.
- `RenderPlan` gives a testable boundary between planning and Excel mutation.

Risks:

- `RenderPlanner` is too large and mixes feature-specific planning for placeholder, row loop, column loop, grid, image, and block behavior.
- `ExcelJsWorksheetRenderer` handles dispatch plus clone row/column logic directly.
- Block rendering is split well, but block planning still lives inside `RenderPlanner`.
- Large workbook behavior is constrained by ExcelJS in-memory workbooks.

## SOLID Review

Single Responsibility:

- Good: `ExpressionEvaluator`, `JsonPathResolver`, `ExcelJsFormulaManager`, and merge classes are focused.
- Needs work: `RenderPlanner` should be decomposed into feature planners such as `PlaceholderPlanner`, `LoopPlanner`, `GridPlanner`, `BlockPlanner`, and `ImagePlanner`.
- Needs work: `ExcelJsWorksheetRenderer` should delegate row/column clone operations to managers the same way block clone does.

Open/Closed:

- Good: new render operations can be added via `RenderOperation`.
- Needs work: operation dispatch currently uses `if` chains. A handler registry would let new operation types plug in without editing `ExcelJsWorksheetRenderer`.

Liskov Substitution:

- Mostly healthy. Ports are narrow and async where needed.
- Future renderers must honor operation ordering within a worksheet.

Interface Segregation:

- Good: manager ports are small.
- Watch: `WorkbookRenderer` combines load/apply/write lifecycle. This is fine for ExcelJS but may be awkward for streaming backends.

Dependency Inversion:

- Good: application depends on ports, infrastructure implements them.
- Fixed in this review pass: removed the type-level circular dependency between `engine/types.ts` and `managers/ports.ts`.

## Clean Architecture Findings

1. Core layer is clean enough for production hardening.
2. Application layer is the right place for `RenderPlan`, but feature planning should be modularized.
3. Infrastructure layer is appropriately ExcelJS-specific.
4. Public API exports many internal classes through `src/index.ts`; before 1.0, decide which exports are stable and which are advanced/internal.

## Memory Leak Review

No obvious event listener, timer, cache, or global retained-state leak was found. The primary memory risk is workload size:

- `ExcelJsWorkbookRenderer.load()` loads the workbook in memory.
- `scanWorkbook()` materializes template source for all worksheets.
- `RenderPlanner.createPlan()` materializes all operations before applying them.
- block clone snapshots materialize all cells in a source block.

These are not leaks, but they are unbounded memory-growth points for very large workbooks.

## Circular Dependency Review

An internal scan originally found:

```text
src/application/engine/types.ts -> src/application/managers/ports.ts -> src/application/engine/types.ts
```

This was removed by defining `TemplateInput` in `src/application/managers/ports.ts` and re-exporting it from `engine/types.ts`.

## Production Readiness Verdict

Status: production-capable for the stated minimal upper-bound benchmarks, but not yet fully proven for complex upper-bound templates.

Ready for:

- small and medium workbook rendering;
- typed helper registration;
- placeholder/default expression rendering;
- row/column/grid/block/image rendering;
- multi-sheet rendering with worksheet-level concurrency.

Proven for minimal stress templates:

- 100k rows;
- 5k columns;
- 50 worksheets;

The planning stage has benchmark coverage for these targets. Full ExcelJS apply/write now completes for minimal 100k-row, 5k-column, and 50-worksheet workbooks. The 100k-row path improved after fixing O(n^2) operation grouping in `ExcelJsWorkbookRenderer.apply()`.

Not yet fully proven for:

- adversarial templates;
- styled/merged/formula-heavy 100k-row workbooks;
- strict package API stability;
- streaming or memory-bounded rendering.
