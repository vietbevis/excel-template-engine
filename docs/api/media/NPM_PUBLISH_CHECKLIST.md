# NPM Publish Checklist

Date: 2026-06-21

## Current Package State

- Package name: `excel-template-engine`
- Version: `0.1.0`
- Module type: ESM
- Entry: `dist/index.js`
- Types: `dist/index.d.ts`
- Included files: `dist`, `README.md`, `LICENSE`, `CHANGELOG.md`, `docs`
- License field: `MIT`
- Node engine: `>=18.18`
- License file: `LICENSE`
- Release notes: `CHANGELOG.md`

## Verified Commands

```bash
npm run typecheck
npm test
npm audit --omit=dev
npm pack --dry-run
```

Current result:

- typecheck: passed
- tests: 56/56 passed
- audit: 0 vulnerabilities
- dry-run package after `build` and Typedoc generation: 172.9 kB, 293 files

## Required Before Publishing

- [x] Add a real `LICENSE` file, because package metadata says MIT.
- [ ] Decide whether `.js.map` and `.d.ts.map` should be published.
- [ ] Decide whether `docs/api` should be generated and included; current dry run includes it.
- [ ] Confirm package name availability on npm.
- [x] Run `npm run docs:api`.
- [x] Run `npm pack --dry-run` after docs generation.
- [ ] Install the packed tarball in a fresh sample project and run a render smoke test.
- [x] Add changelog or release notes.
- [ ] Define support matrix for Node versions and ExcelJS version.
- [x] Add API stability statement for `0.1.x`.
- [ ] Add provenance/2FA release process for npm publishing.

## Suggested Release Script

```bash
npm run clean
npm run build
npm run typecheck
npm test
npm run docs:api
npm run benchmark:large
npm audit --omit=dev
npm pack --dry-run
```

## Do Not Publish Until

- stress targets are documented with the latest benchmark numbers;
- trusted/untrusted template model is documented;
- asset path security policy stays covered by tests and docs;
- fresh tarball install smoke test is run before the actual publish.
