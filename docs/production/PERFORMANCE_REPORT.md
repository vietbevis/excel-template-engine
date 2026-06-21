# Performance Report

Date: 2026-06-21

## Executive Summary

Current benchmarks show the core placeholder path is fast and style cloning is acceptable for small batches. However, the repo does not yet contain stress benchmarks for 100k rows, 5k columns, or 50 worksheets. The implementation is still in-memory through ExcelJS, so production claims for very large files need stress tests and likely architecture changes.

## Commands Run

```bash
npm run benchmark:placeholder
npm run benchmark:style-clone
npm test
```

## Current Results

Placeholder benchmark:

```json
{
  "iterations": 50000,
  "evaluator": {
    "durationMs": 38,
    "opsPerSecond": 1315495
  },
  "placeholderRender": {
    "durationMs": 138,
    "opsPerSecond": 361715
  }
}
```

Style clone benchmark:

```json
{
  "iterations": 5000,
  "cell": {
    "durationMs": 25,
    "opsPerSecond": 199715
  },
  "row": {
    "durationMs": 26,
    "opsPerSecond": 189387
  },
  "column": {
    "durationMs": 1711,
    "opsPerSecond": 2922
  }
}
```

Test suite:

```text
51/51 tests passed
```

## Scale Target Assessment

| Target | Current evidence | Status |
| --- | --- | --- |
| 100k rows | No stress benchmark yet. ExcelJS workbook is loaded and written in memory. | Not proven |
| 5k columns | No stress benchmark yet. Column clone copies styles over worksheet row count. | Not proven |
| 50 worksheets | Multi-sheet integration exists, but only small sheets. | Partially proven |
| Cross-sheet formula | Integration test exists. | Proven for simple formulas |
| Cross-sheet block | Integration test exists. | Proven for simple block |

## Hot Paths

1. `RenderPlanner.createPlan()` walks every scanned cell and accumulates all operations.
2. `ExcelJsWorkbookRenderer.scanWorkbook()` materializes workbook template source.
3. `ExcelJsWorksheetRenderer.applyCloneColumn()` loops all rows for each cloned column.
4. `ExcelJsWorksheetRenderer.applyCloneRow()` loops all columns for each cloned row.
5. `ExcelJsBlockCloneManager.createSnapshots()` snapshots every cell in a block range.
6. `ExcelJsMergeManager.collectMergeRanges()` reads private ExcelJS `_merges` state repeatedly.

## Risks

- For 100k rows, operation arrays can become very large before any rendering happens.
- For 5k columns, row/column clone loops become expensive and may clone styles for empty cells.
- For 50 worksheets, current worksheet-level `Promise.all` can increase peak memory if many large sheets render at once.
- `writeBuffer()` creates a full in-memory output buffer.

## Recommendations

1. Add stress benchmarks:
   - `benchmark:large-rows` for 100k row placeholder/each rendering.
   - `benchmark:large-columns` for 5k column each-col rendering.
   - `benchmark:many-sheets` for 50 worksheets with mixed placeholders and formulas.
2. Add memory instrumentation using `process.memoryUsage()` before load, after scan, after plan, after apply, and after write.
3. Replace global operation accumulation with per-sheet or streaming planning.
4. Limit worksheet concurrency with a configurable `maxConcurrentWorksheets`.
5. Optimize style cloning to skip empty/default cells.
6. Cache merge collection during a sheet operation batch.
7. Consider an XLSX XML-level renderer for very large clone workloads where ExcelJS in-memory mutation is too expensive.

## Production SLO Proposal

For a first production release, define measured limits instead of aspirational limits:

- maximum tested worksheets;
- maximum tested rows per worksheet;
- maximum tested columns per worksheet;
- maximum output size;
- maximum image payload size;
- expected memory ceiling on Node 18 and Node 20.
