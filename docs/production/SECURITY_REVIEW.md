# Security Review

Date: 2026-06-21

## Executive Summary

No dependency vulnerabilities were reported by `npm audit --omit=dev`. The project avoids `eval` and the `Function` constructor in expression evaluation, which is a strong default for template safety. The main security concerns are filesystem asset loading, unbounded input sizes, formula injection semantics, and lack of documented trust boundaries.

This is a TypeScript/Node library, not a web server. The available security skill references were web-framework-specific and did not directly apply, so this review uses general Node package security practices plus source inspection.

## Evidence Reviewed

- `npm audit --omit=dev`: 0 vulnerabilities.
- `ExpressionEvaluator` parses a restricted expression language instead of executing JavaScript.
- `DefaultAssetResolver` reads user-provided image paths from disk.
- Excel formula cells are preserved/shifted and may execute when opened by Excel.

## Findings

### High

S1. Unrestricted image path reads

Evidence: `src/infrastructure/assets/default-asset-resolver.ts:39-40` resolves a string source and reads it from disk.

Impact: If untrusted JSON controls `{{image avatar}}`, it can cause the library host process to read arbitrary local files that exist and look like PNG/JPG.

Recommendation:

- Add an `AssetResolverOptions` policy with `allowAbsolutePaths`, `baseDir`, and `maxBytes`.
- Default to rejecting absolute paths unless explicitly enabled.
- Reject path traversal outside `baseDir`.

S2. No input size limits

Evidence: `src/infrastructure/assets/default-asset-resolver.ts:31-40` decodes base64 or reads whole files; `src/infrastructure/exceljs/excel-js-workbook-renderer.ts:20-31` loads whole templates; `write()` returns a full `Uint8Array`.

Impact: Untrusted templates or image payloads can cause memory exhaustion.

Recommendation:

- Add configurable limits for template bytes, image bytes, worksheets, rows, columns, and render operations.
- Return structured errors when limits are exceeded.

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
- Bounded asset resolver.
- Optional formula stripping.
- Limit-aware render options.
- Document trusted/untrusted input model.
