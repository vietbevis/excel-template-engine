# NPM Publish Checklist

Date: 2026-06-21

## Current Package State

- Package name: `excel-template-engine`
- Version: `0.1.0`
- Module type: ESM
- Entry: `dist/index.js`
- Types: `dist/index.d.ts`
- Included files: `dist`, `README.md`, `docs`
- License field: `MIT`
- Node engine: `>=18.18`

## Verified Commands

```bash
npm run typecheck
npm test
npm audit --omit=dev
npm pack --dry-run
```

Current result:

- typecheck: passed
- tests: 51/51 passed
- audit: 0 vulnerabilities
- dry-run package: 59.1 kB, 170 files

## Required Before Publishing

- [ ] Add a real `LICENSE` file, because package metadata says MIT.
- [ ] Decide whether `.js.map` and `.d.ts.map` should be published.
- [ ] Decide whether `docs/api` should be generated and included.
- [ ] Confirm package name availability on npm.
- [ ] Run `npm run docs:api`.
- [ ] Run `npm pack --dry-run` after docs generation.
- [ ] Install the packed tarball in a fresh sample project and run a render smoke test.
- [ ] Add changelog or release notes.
- [ ] Define support matrix for Node versions and ExcelJS version.
- [ ] Add provenance/2FA release process for npm publishing.

## Suggested Release Script

```bash
npm run clean
npm run build
npm run typecheck
npm test
npm run docs:api
npm audit --omit=dev
npm pack --dry-run
```

## Do Not Publish Until

- stress targets are either proven or documented as unsupported;
- trusted/untrusted template model is documented;
- asset path security policy is implemented or documented as host responsibility;
- API stability statement is written.
