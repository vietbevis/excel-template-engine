# Refactor Plan

Date: 2026-06-21

## Priority 0: Keep Current Gates Green

- Keep `npm run typecheck` passing.
- Keep `npm test` passing.
- Keep `npm audit --omit=dev` clean.
- Add stress benchmarks before claiming high-scale readiness.

## Priority 1: Scale and Safety

1. Add `RenderLimits`:
   - `maxWorksheets`
   - `maxRows`
   - `maxColumns`
   - `maxOperations`
   - `maxImageBytes`
   - `maxTemplateBytes`
2. Add `AssetResolverOptions`:
   - reject absolute paths by default;
   - enforce baseDir containment;
   - enforce max byte size.
3. Add `FormulaPolicy`:
   - `preserve`
   - `strip`
   - `shift`
4. Add stress benchmarks for:
   - 100k rows;
   - 5k columns;
   - 50 worksheets.

## Priority 2: Decompose Planning

Split `RenderPlanner` into feature-specific planners:

- `PlaceholderPlanner`
- `EachRowPlanner`
- `EachColumnPlanner`
- `GridPlanner`
- `BlockPlanner`
- `ImagePlanner`

Each planner should return `RenderOperation[]` and warnings. This will reduce the current large-file risk in `render-planner.ts` and make new syntax easier to add.

## Priority 3: Renderer Operation Handlers

Replace `ExcelJsWorksheetRenderer.apply()` `if` dispatch with an operation-handler registry:

- `SetCellValueHandler`
- `CloneRowHandler`
- `CloneColumnHandler`
- `CloneBlockHandler`
- `InsertImageHandler`

This improves open/closed design and makes operation-specific tests smaller.

## Priority 4: Performance Architecture

1. Plan and apply per worksheet instead of creating a workbook-wide operation list.
2. Add configurable worksheet concurrency.
3. Cache merge ranges per sheet operation batch.
4. Skip style clone for empty/default cells.
5. Investigate XML-level clone backend for massive row/column expansion.

## Priority 5: Public API Stabilization

1. Split exports into stable API and advanced internals.
2. Add semantic versioning rules.
3. Document migration rules for AST/render-plan changes.
4. Decide whether managers are public extension points or internal implementation details.

## Priority 6: xlsx-template Differentiators

To exceed `xlsx-template`, focus on:

- typed API and TypeScript declarations;
- safe expression evaluator without `eval`;
- render-plan inspection/debugging;
- helper registry with typed context;
- multi-sheet rendering;
- image/base64/buffer support;
- formula shift manager;
- explicit production limits and benchmark reports;
- pluggable renderer/asset resolver architecture.
