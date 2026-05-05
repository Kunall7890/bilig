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
