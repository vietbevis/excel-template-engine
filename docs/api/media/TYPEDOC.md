# Typedoc

Date: 2026-06-21

## Status

Typedoc is configured through:

```text
typedoc.json
```

Command:

```bash
npm run docs:api
```

Entry point:

```text
src/index.ts
```

Output:

```text
docs/api
```

Latest verification:

```bash
npm run docs:api
```

Result: generated HTML at `docs/api` without warnings.

## Publishing Policy

Before publishing, decide whether generated `docs/api` should be committed and shipped in the npm package. The current package `files` field includes `docs`, so generated API docs will be included if present.

## Recommended CI Gate

```bash
npm run docs:api
```

The command should be part of release verification once API docs are considered required output.
