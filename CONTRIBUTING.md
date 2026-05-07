# Contributing to bilig

Thanks for taking the repo seriously. bilig is an engine-heavy TypeScript
monorepo, so the best contributions are small, tested, and explicit about which
runtime behavior they change.

## Local Setup

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```bash
pnpm install
pnpm wasm:build
pnpm typecheck
pnpm test
```

For the app shell:

```bash
pnpm dev:web-local
```

## Before You Open a PR

Run the narrowest checks that cover your change, then run the full gate when the
change is ready to publish.

```bash
pnpm lint
pnpm typecheck
pnpm test
pnpm test:browser
pnpm run ci
```

If you edit generated protocol, formula inventory, workspace-resolution, or
benchmark-baseline sources, regenerate and commit the generated output.

## Good First Areas

- Add formula fixtures and tests for missing Excel-compatible semantics.
- Turn architecture docs into smaller runnable examples.
- Improve grid accessibility, keyboard behavior, and focus handling.
- Add WorkPaper benchmark cases that describe the real spreadsheet pattern they
  represent.
- Tighten engine correctness tests around mutation, snapshot, undo/redo, and
  dependency behavior.

## Formula Parity Fixture Walkthrough

Start by reading the existing fixture shape before adding a new case:

- `packages/excel-fixtures/src/` is the canonical formula and workbook-semantics
  corpus used by `@bilig/formula` and engine runtime checks.
- `packages/formula/src/compatibility.ts` records implementation status and
  generated inventory metadata for formula families.
- `packages/formula/src/__tests__/fixture-harness.test.ts` executes implemented
  canonical fixtures through the JavaScript evaluator.
- `packages/core/src/__tests__/formula-runtime-correctness.test.ts` covers the
  production runtime path for fixtures that should run through the engine.
- `packages/headless/fixtures/xlsx-corpus/` holds checked-in XLSX cached-result
  reductions for public workbook compatibility regressions.

A minimal formula-parity contribution should:

1. Add or tighten one small fixture with an Excel-observed expected result.
2. Update the compatibility entry only to the status that the implementation
   actually supports.
3. Add focused package tests when the fixture exposes behavior not already
   covered by the harness.
4. Run the generated checks before opening a PR:

```bash
pnpm formula-inventory:check
pnpm formula:dominance:check
pnpm test:correctness:formula
```

For cached-result workbook reductions, regenerate or verify the headless XLSX
fixture corpus instead of hand-editing binary evidence:

```bash
pnpm workpaper:xlsx-corpus:fixtures:generate
pnpm workpaper:xlsx-corpus:fixtures:check
```

Before claiming the work is ready, run `pnpm run ci`. Do not describe a fixture
as full Excel parity; the checked-in fixtures prove only the named formulas,
inputs, cached workbook results, and runtime paths they actually cover.

## Contribution Rules

- Keep public APIs boring and stable. Prefer `is...`, `allows...`, `on...`, and
  `on...Change` naming.
- Keep formula semantics in JavaScript first. Promote to WASM only after parity
  and differential tests are green.
- Avoid `any`; lint fails on weak typing and floating promises.
- Use explicit `.js` suffixes where nearby ESM imports already do.
- Do not mix UI rendering, behavior policy, and workbook engine logic in one
  component when a hook, controller, or package boundary can own one concern.
- Keep benchmark claims tied to commands, artifacts, counters, or checked-in
  fixtures.

## PR Description

Include:

- what changed
- why the change belongs in this package
- commands run
- benchmark output or screenshots when behavior is visual or performance-related
- known risk or follow-up work

## Source of Truth

Forgejo `origin` is the primary repo workflow for maintainers. GitHub mirrors
the public verification contract and public collaboration surface.
