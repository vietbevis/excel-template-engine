# Performance Report

Date: 2026-06-21

## Executive Summary

Current benchmarks show the core placeholder path is fast, style cloning is acceptable for small batches, and planning handles 100k row expansion, 5k column expansion, and 50 worksheets quickly. Full ExcelJS apply/write is now proven for minimal 100k-row, 5k-column, and 50-worksheet workbooks after removing an O(n^2) operation grouping bottleneck in `ExcelJsWorkbookRenderer.apply()`.

## Commands Run

```bash
npm run benchmark:placeholder
npm run benchmark:style-clone
npm run benchmark:large
npm run benchmark:large-render:rows
npm run benchmark:large-render:columns
npm run benchmark:large-render:sheets
npm test
```

Large planning benchmark:

```json
{
  "rows100k": {
    "name": "plan_100k_rows_each",
    "durationMs": 70,
    "operations": 100001,
    "operationsPerSecond": 1433275,
    "memory": {
      "heapUsedDeltaMb": 55.43,
      "rssDeltaMb": 103.61,
      "heapUsedMb": 67.54,
      "rssMb": 164.17
    }
  },
  "columns5k": {
    "name": "plan_5k_columns_each_col",
    "durationMs": 4,
    "operations": 5001,
    "operationsPerSecond": 1365771,
    "memory": {
      "heapUsedDeltaMb": -14.45,
      "rssDeltaMb": 1.28,
      "heapUsedMb": 53.09,
      "rssMb": 165.47
    }
  },
  "worksheets50": {
    "name": "plan_50_worksheets",
    "durationMs": 0,
    "operations": 50,
    "operationsPerSecond": 147113,
    "memory": {
      "heapUsedDeltaMb": 0.57,
      "rssDeltaMb": 0.03,
      "heapUsedMb": 53.66,
      "rssMb": 165.5
    }
  }
}
```

Full ExcelJS render/write benchmark:

```json
{
  "rows": {
    "name": "render_100k_rows_each",
    "durationMs": 715,
    "outputBytes": 1097796,
    "memory": {
      "heapUsedDeltaMb": 106.78,
      "rssDeltaMb": 342.42,
      "heapUsedMb": 120.98,
      "rssMb": 414.09
    }
  },
  "columns": {
    "name": "render_5k_columns_each_col",
    "durationMs": 54,
    "outputBytes": 43175,
    "memory": {
      "heapUsedDeltaMb": 16.65,
      "rssDeltaMb": 52.27,
      "heapUsedMb": 30.78,
      "rssMb": 123.58
    }
  },
  "sheets": {
    "name": "render_50_worksheets",
    "durationMs": 54,
    "outputBytes": 36647,
    "memory": {
      "heapUsedDeltaMb": 4.26,
      "rssDeltaMb": 94.72,
      "heapUsedMb": 18.3,
      "rssMb": 165.75
    }
  }
}
```

The 100k-row full render previously timed out because `ExcelJsWorkbookRenderer.apply()` copied the accumulated sheet operation array on every operation. The renderer now groups operations with append-only arrays and applies synchronous `SetCellValue` operations without one `await` per cell.

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
56/56 tests passed
```

## Scale Target Assessment

| Target | Current evidence | Status |
| --- | --- | --- |
| 100k rows | Planning benchmark creates 100001 operations in 70 ms. Full ExcelJS render/write completes in 715 ms with 1.10 MB output. | Proven for minimal template |
| 5k columns | Planning benchmark creates 5001 operations in 4 ms. Full ExcelJS render/write completes in 54 ms. | Proven for minimal template |
| 50 worksheets | Planning benchmark and full ExcelJS render/write complete; full render/write completes in 54 ms. | Proven for minimal template |
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

- For 100k rows, operation arrays can become very large before any rendering happens, even though the current minimal render benchmark passes.
- For 5k columns, row/column clone loops become expensive and may clone styles for empty cells.
- For 50 worksheets, current worksheet-level `Promise.all` can increase peak memory if many large sheets render at once.
- `writeBuffer()` creates a full in-memory output buffer.

## Recommendations

1. Keep full ExcelJS render/write benchmarks for 100k rows, 5k columns, and 50 worksheets in CI or release checks.
2. Extend memory instrumentation from planning to load, scan, apply, and write.
3. Replace global operation accumulation with per-sheet or streaming planning for lower peak memory.
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
