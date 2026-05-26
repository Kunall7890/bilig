# Fix Analyze Quality Dead Exports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Clear the `pnpm analyze:quality` and `pnpm test:correctness:fast` dead-export failure from audit item 3 without changing workbook-agent runtime behavior.

**Architecture:** The two flagged helpers are local implementation details, not public module API. Keep their call sites unchanged and remove only the `export` modifiers so `knip --include exports` stops treating them as unused public exports.

**Tech Stack:** TypeScript, ESM, pnpm, knip, Vitest-backed correctness scripts.

---

### Task 1: Confirm the Analyzer Failure

**Files:**
- Inspect: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Inspect: `apps/bilig/src/codex-app/workbook-agent-visible-commit-barrier.ts`

- [x] **Step 1: Run the failing analyzer**

Run:

```bash
pnpm analyze:exports
```

Expected: FAIL with these two unused exports:

```text
selectWorkbookAgentRenderedVerificationRanges
summarizeWorkbookAgentVisibleCommitBarrierOutcome
```

- [x] **Step 2: Confirm usage is module-local**

Run:

```bash
rg -n "selectWorkbookAgentRenderedVerificationRanges|summarizeWorkbookAgentVisibleCommitBarrierOutcome" apps/bilig/src packages apps/web scripts docs
```

Expected: each function appears only in its defining file and its same-file call site.

### Task 2: Make the Helpers Private

**Files:**
- Modify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Modify: `apps/bilig/src/codex-app/workbook-agent-visible-commit-barrier.ts`

- [x] **Step 1: Remove the public export from the rendered-verification selector**

Change:

```ts
export function selectWorkbookAgentRenderedVerificationRanges(
```

to:

```ts
function selectWorkbookAgentRenderedVerificationRanges(
```

- [x] **Step 2: Remove the public export from the visible-commit summary helper**

Change:

```ts
export function summarizeWorkbookAgentVisibleCommitBarrierOutcome(input: {
```

to:

```ts
function summarizeWorkbookAgentVisibleCommitBarrierOutcome(input: {
```

### Task 3: Verify the Gate

**Files:**
- Verify: `apps/bilig/src/codex-app/workbook-agent-dynamic-tool-handler.ts`
- Verify: `apps/bilig/src/codex-app/workbook-agent-visible-commit-barrier.ts`

- [x] **Step 1: Run focused export analysis**

Run:

```bash
pnpm analyze:exports
```

Expected: PASS.

- [x] **Step 2: Run the full analyzer quality aggregate**

Run:

```bash
pnpm analyze:quality
```

Expected: PASS.

- [x] **Step 3: Run the dependent fast correctness gate**

Run:

```bash
pnpm test:correctness:fast
```

Expected: PASS.

- [x] **Step 4: Run TypeScript validation for the touched app boundary**

Run:

```bash
pnpm typecheck
```

Expected: PASS.
