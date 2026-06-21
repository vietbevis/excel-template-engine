# Proposal: Exceeding xlsx-template

Date: 2026-06-21

## Current Advantages

- TypeScript-first public API and emitted declarations.
- Restricted expression evaluator without JavaScript execution.
- Helper registry with context.
- Multi-sheet rendering.
- Image source support for paths, buffers, and base64.
- Formula shifting for cloned blocks.
- Explicit managers for style, merge, formula, image, and asset resolution.
- Render-plan architecture that can support debugging and dry-run inspection.

## Gaps to Close

- Prove large workbook performance.
- Add limit policies for untrusted data/templates.
- Add a streaming or XML-level backend for massive clone operations.
- Add better formula parser coverage for advanced Excel syntax.
- Add a template diagnostics mode that lists all placeholders/helpers/images before rendering.

## Differentiators to Build

1. `engine.analyze(template)`:
   - list placeholders;
   - list helpers;
   - list image sources;
   - list blocks/loops/grids;
   - report missing data paths before rendering.
2. `engine.plan(template, data)` public inspection:
   - return `RenderPlan`;
   - estimate operations and output size risk.
3. Production limits:
   - fail fast on huge sheets/images/operation counts.
4. Formula policy:
   - preserve, shift, strip, or recalculate metadata.
5. Template validation CLI:
   - useful in CI for template authors.
6. Better benchmark story:
   - publish repeatable benchmark fixtures.
7. Renderer adapters:
   - ExcelJS for compatibility;
   - XML-level renderer for very large workbooks.
