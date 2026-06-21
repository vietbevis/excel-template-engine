# Security Review

Date: 2026-06-21

## Executive Summary

No dependency vulnerabilities were reported by `npm audit --omit=dev`. The project avoids `eval` and the `Function` constructor in expression evaluation, which is a strong default for template safety. Filesystem asset loading and image size limits now have safer defaults. The main remaining security concerns are template/output size limits, formula trust boundaries, and private ExcelJS API compatibility.

This is a TypeScript/Node library, not a web server. The available security skill references were web-framework-specific and did not directly apply, so this review uses general Node package security practices plus source inspection.

## Evidence Reviewed

- `npm audit --omit=dev`: 0 vulnerabilities.
- `ExpressionEvaluator` parses a restricted expression language instead of executing JavaScript.
- `DefaultAssetResolver` resolves relative image paths inside a base directory, rejects absolute paths by default, and enforces image byte limits.
- Excel formula cells are preserved/shifted and may execute when opened by Excel.

## Findings

### High

S1. Unrestricted image path reads - mitigated by default

Evidence: `DefaultAssetResolver` resolves relative paths inside a configured `baseDir`, rejects absolute image paths by default, and rejects path traversal outside `baseDir`.

Residual impact: hosts can explicitly enable absolute paths with `allowAbsolutePaths`, so production apps must only enable that for trusted data.

Recommendation:

- Keep `allowAbsolutePaths` disabled for untrusted data.
- Document the trust boundary for custom `AssetResolver` implementations.

S2. No input size limits - partially mitigated

Evidence: `DefaultAssetResolver` has `maxBytes` for file, buffer, and base64 images. `RenderPlanner` has `limits` for worksheets, rows, columns, and render operations. `ExcelJsWorkbookRenderer` enforces `maxTemplateBytes` before loading template files or buffers.

Residual impact: full ExcelJS output size is still not bounded by default.

Recommendation:

- Add output/write memory guidance.
- Continue expanding full render benchmarks.
- Structured `LimitExceededError` is now available for limit violations.

### Medium

S3. Formula trust boundary is not documented

Evidence: formulas are preserved and shifted in `src/infrastructure/exceljs/excel-js-block-clone-manager.ts:163-178`.

Impact: XLSX consumers may execute formulas when opening generated files. If template authors are untrusted, this can create spreadsheet formula risk.

Recommendation:

- Document that templates are trusted input by default.
- Add an option to strip formulas or force plain values for untrusted templates.

S4. Internal ExcelJS private API use

Evidence: `ExcelJsMergeManager` reads `_merges`, a private ExcelJS field.

Impact: Dependency updates can break merge handling in ways that may corrupt output.

Recommendation:

- Pin ExcelJS minor version or add compatibility tests for supported ExcelJS versions.
- Isolate private API access behind a single adapter with tests.

### Low

S5. Package publishes source maps

Evidence: `npm pack --dry-run` includes `.js.map` and `.d.ts.map` files.

Impact: Source maps are usually fine for libraries, but they increase published surface and reveal source layout.

Recommendation:

- Decide whether source maps are intentional. If not, exclude them or disable declaration/source maps for published builds.

## Secure Defaults Proposed

- No JavaScript evaluation in templates.
- Explicit missing-value policy.
- Bounded asset resolver with default absolute-path rejection.
- Optional formula stripping.
- Limit-aware render options.
- Document trusted/untrusted input model.
