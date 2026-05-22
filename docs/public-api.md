# Public APIs

## Current state

- `@bilig/core` now implements the documented range-mutation, undo/redo, selection, and sync-state surface.
- `SelectionState` is additive: existing callers can keep using `sheetName` and `address`, while newer callers can use `anchorAddress`, `range`, and `editMode`.
- `connectSyncClient` is live. `apps/web` is already worker-first by default, but the remote sync service is not a closed durable worksheet backend yet.
- `@bilig/binary-protocol` is already a real wire protocol for sync frames, but the authoritative replicated workbook op family is still narrower than the full local engine surface.
- `@bilig/agent-api` is currently a binary envelope around JSON payloads, not yet a fully typed binary request/response/event schema.

## Stable packages

`@bilig/headless` is the current external npm adoption path for WorkPaper
calculation and Excel workbook import/export. Its `@bilig/headless/xlsx`
subpath exposes the repository XLSX importer/exporter without requiring a second
public npm package. The other package surfaces listed here are stable repository
package boundaries; not every package name is provisioned on npm yet.

- `@bilig/core`
- `@bilig/headless`
- `@bilig/excel-import`
- `@bilig/formula`
- `@bilig/wasm-kernel`
- `@bilig/workbook`
- `@bilig/renderer`
- `@bilig/grid`
- `@bilig/binary-protocol`
- `@bilig/worker-transport`
- `@bilig/agent-api`
- `@bilig/storage-server`
- `@bilig/excel-fixtures`

## Workbook DSL

`@bilig/renderer` keeps the declarative workbook DSL unchanged:

- `<Workbook>`
- `<Sheet name="...">`
- `<Cell addr="..." value={...} />`
- `<Cell addr="..." formula="..." />`
- `<Cell addr="..." format="..." />`

## Agent-first workbook surface

`@bilig/workbook` is the generic public package for consumer-defined workbook
models. It does not ship business-model templates and does not depend on
`@bilig/core`, `@bilig/headless`, `@bilig/agent-api`, `zod`, or `effect`.

Build `@bilig/workbook` so an agent would love using it: simple, generic,
predictable, inspectable, verifiable, and never dependent on hardcoded business
models or human spreadsheet UI assumptions.

It exposes:

- `defineModel`
- `buildWorkbookActionPlan`
- `planWorkbookAction`
- `inspectModel`
- `collectWorkbookRefs`
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows`
- `find`
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRunResult`
- `describeRuntimeRequirements`
- `verifyPlan`
- `verifyModel`
- `runWorkbookPlan`
- `runWorkbookAction`
- `verifyWorkbookReadbacks`
- `normalizeWorkbookActionInputDescription`
- `workbookPlanIssueCodes`
- `isWorkbookPlanIssueCode`
- `workbookReadbackIssueCodes`
- `isWorkbookReadbackIssueCode`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` through the model workbook context and as top-level helpers
- `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.valuesEqual`, `check.formulaEquals`, `check.formulasEqual`, and `check.custom` through the model workbook context and as top-level helpers
- `WorkbookModel`, `WorkbookAction`, `WorkbookActionConfig`, `WorkbookActionDefinition`, `WorkbookActionContext`, `WorkbookCheckContext`, `WorkbookFindWorkbook`, `WorkbookCheckWorkbook`, `WorkbookActionWorkbook`, `WorkbookModelWorkbook`, `WorkbookFindNamespace`, `WorkbookActionInput`, `WorkbookActionInputDescription`, `WorkbookActionInputDescriptionKind`, `WorkbookActionInspection`, `WorkbookAddOpOptions`, `WorkbookActionPlanResult`, `WorkbookModelDescription`, `WorkbookRefDescription`, `WorkbookActionPlanDescription`, `WorkbookActionPlanResultDescription`, `WorkbookRunResultDescription`, `WorkbookUndoRefDescription`, `WorkbookAppliedSummaryDescription`, `WorkbookRuntimeRequirements`, `WorkbookRuntimeRequirement`, `WorkbookRuntimeCapability`, `WorkbookRuntimeMaterialization`, `WorkbookRuntimePreview`, `WorkbookPlanVerification`, `WorkbookPlanIssue`, `WorkbookPlanIssueCode`, `WorkbookModelVerification`, `WorkbookModelActionVerification`, `WorkbookModelVerificationOptions`, `WorkbookRunAdapter`, `WorkbookRunApplyResult`, `WorkbookCellReadback`, `WorkbookRunReadback`, `WorkbookReadbackVerification`, `WorkbookReadbackIssue`, `WorkbookReadbackIssueCode`, `WorkbookCheckExpectation`, `WorkbookCheckExpectationDescription`, `WorkbookCustomCheckOptions`, `WorkbookReadbackCheckOptions`, `WorkbookRawFormulaOptions`, `WorkbookRunResult`, `WorkbookAppliedSummary`, `WorkbookRunError`, `WorkbookRunErrorCode`, and `WorkbookCheckResult`
- the existing low-level operation language: `WorkbookOp`, `WorkbookTxn`, `EngineOp`, and `EngineOpBatch`

The package builds portable workbook intent and concrete low-level ops when the
target is already known. Formula helpers use `@bilig/formula` for parsing and
normalization. Actual calculation and authoritative execution stay in
`@bilig/core` and `apps/bilig`.

Model actions can accept plain JSON-safe input through
`planWorkbookAction(model, actionName, input)` and
`buildWorkbookActionPlan(model, actionName, input)`. The input is cloned,
canonicalized with stable object-key order, recorded on the plan, and exposed to
the action context as `input`. Supported values are strings, finite numbers,
booleans, `null`, arrays without holes, and plain objects. `@bilig/workbook`
validates action-object input metadata before model code runs, using only simple
JSON kinds, required fields, and array item kinds. It still does not provide a
schema framework for consumer business meaning. `verifyModel(model, { inputs })`
supplies per-action inputs for whole-model verification.

When agents need to know what an action expects before running workbook code,
actions may be declared as action objects:

```ts
actions: {
  write: {
    description: 'Write a consumer-provided value',
    input: {
      kind: 'object',
      fields: {
        value: { kind: 'number', required: true },
      },
    },
    run({ refs, workbook, input }) {
      if (typeof input !== 'object' || input === null || Array.isArray(input)) {
        throw new Error('input object required')
      }
      const value = input.value
      if (typeof value !== 'number') {
        throw new Error('numeric value required')
      }
      workbook.writeValue(refs.output, value)
    },
  },
}
```

This is descriptive metadata, not a schema framework. Input descriptions use
plain JSON kinds: `json`, `object`, `array`, `string`, `number`, `boolean`, and
`null`. `object` descriptions may list `fields`; `array` descriptions may list
`items`; fields and root inputs may be `required`.
`normalizeWorkbookActionInputDescription` trims text, rejects malformed
metadata, freezes the result, and keeps `@bilig/workbook` independent from
`zod`, `effect`, and model-specific validators.

Known single-cell `workbook.format(ref, { numberFormat })` actions compile to
concrete `setCellFormat` ops, including `numberFormat: null` for explicit
format clears. Style patches remain high-level intent until the runtime resolves
style ids.
When a consumer needs a workbook operation that is already covered by the
low-level operation language, model actions can call
`workbook.addOp(op, { target?, message? })`. The op is guarded with
`isWorkbookOp`, cloned into `plan.ops`, and kept as a command so agents can
inspect the handoff without pulling in `@bilig/core`. When a `target` is
supplied for an address or range op, `verifyPlan` checks that the op touches the
same range. For op kinds without an inferable range, `target` is descriptive for
logs and approvals rather than proof of affected cells.
The public op guards validate finite literal values, parseable formulas, valid
cell addresses, ordered ranges, non-empty identifiers, and known enum values
before a low-level op is accepted.

Formula helpers keep referenced workbook inputs separate from formula text.
Planned `writeFormula` commands expose those inputs directly, which lets agents
inspect dependencies without relying on human UI coordinates or reverse-parsing
placeholder formula names.
For formulas outside the small helper set, `formula.raw(source, { inputs })`
keeps arbitrary formula text generic while preserving explicit workbook
dependencies for inspection and verification. These are declared dependencies,
not parser-discovered proof that every formula reference has a matching model
ref.
Runtime adapters materialize declared formula inputs as whole formula tokens.
They do not rewrite a ref token inside a string literal or inside a larger
identifier, so generic formula text remains predictable for agents.

Action plans also expose `refsUsed`, a flat deduped list of workbook refs found
inside the consumer-defined `refs` object. This keeps custom models generic
while still letting agents inspect what the model resolved.
The same generic refs are available outside model callbacks through top-level
`findTable`, `findColumn`, `findRange`, `findName`, and `findRows` helpers, or
through the frozen `find` namespace with short aliases such as
`find.table(...)`, `find.range(...)`, and `find.rows(...)`.
These selector helpers trim text, canonicalize cell addresses, and reject empty
or malformed selectors before the runtime handoff. That keeps bad agent intent
out of the plan instead of letting an invalid address, blank column, invalid row
operator, or non-finite predicate value fail later inside an engine adapter.
`findRows` refs include their predicate value in the stable id, so distinct
consumer-defined row predicates remain distinct during agent inspection and
dedupe. `findRows` is table-backed: pass the table plus a predicate, then use
`rows.column("Amount")` for row-scoped columns. Ref ids are collision-resistant
and stable but opaque; display labels or `describeRef` output instead of
parsing ids.
Refs are frozen data objects. Helper methods such as `table.column()` and
`rows.column()` remain available for ergonomics, but they are non-enumerable so
object-key inspection and JSON descriptions stay data-first.
For table-backed row selectors, `rows.column("Amount")` targets that column only
inside the matching rows. Core/app runtime adapters can resolve those generic
refs into exact cells for writes, formats, clears, checks, and row-wise formula
input alignment without adding hardcoded workbook models.
The same planned checks are available outside model callbacks through top-level
`check.exists(ref)`, `check.noFormulaErrors(ref)`,
`check.valueEquals(ref, value)`, `check.valuesEqual(ref, rows)`,
`check.formulaEquals(ref, formula)`, `check.formulasEqual(ref, rows)`, and
`check.custom({ kind, message, target, refs })` helpers. Custom checks let
consumers carry their own invariants without adding hardcoded business models to
the package. `target` names the main ref, and `refs` names supporting refs so
agents can describe and verify the full invariant contract.
Readback checks add machine-readable expectations to the same generic check
channel: `valueEquals` and `formulaEquals` cover scalar cells, while
`valuesEqual` and `formulasEqual` cover generic multi-cell targets such as
ranges, table columns, and row-filtered columns. Runtime code can evaluate those
expectations after applying the plan, while agents can inspect the proof target
without relying on visual spreadsheet state.

Model callback phases expose only the API needed for that phase. `find` gets
find helpers only, `checks` gets find helpers and planned-check helpers, and
actions get find helpers, checks, and mutation-planning methods. Discovery and
proof declaration cannot accidentally write workbook intent before the action
phase.

`describeModel` returns a JSON-safe model manifest with the model name, optional
model description, sorted action names, per-action descriptions, optional input
descriptions, and whether model-level checks exist. It does not run `find`,
checks, or actions.
For agent logs, approvals, tests, and runtime handoff, `describeRef` and
`describePlan` produce JSON-safe descriptions of refs and action plans. The
descriptions preserve generic action input and workbook intent while removing
consumer-private `refs` object shape and helper methods.
`describePlanResult` applies the same description layer to either planned or
failed action planning results.
`describeRunResult` applies the same JSON-safe description layer after
execution, preserving `done`/`failed` status, changed summaries, checks, errors,
and undo ops while removing ref helper functions from the public result.
`describeRuntimeRequirements(plan)` gives agents a JSON-safe adapter checklist
for the same plan: which generic commands must be applied, which readbacks are
needed, and which checks need proof. It stays generic, with capabilities such
as `writeFormula`, `writeValue`, `format`, `clear`, `applyOp`, `read`, and
`verifyCheck`. Apply requirements include a stable `path` and a materialization
classification: `concreteOp`, `providedOp`, or `adapterMaterialization`. Read
requirements include the machine-readable expectation. This lets agents see
whether the handoff is already backed by portable ops or needs runtime
materialization without importing the engine.

`verifyPlan` gives agents a runtime-free consistency check before handoff. It
flags invalid action input, unresolved command targets, unresolved formula
inputs, duplicate resolved refs, unparsable formulas, and missing concrete ops
for write, clear, and number-format commands whose target is already known as a
single cell. Custom check targets and supporting refs must also resolve through
`refsUsed`.
Planning issues expose a stable `WorkbookPlanIssueCode` union plus the frozen
`workbookPlanIssueCodes` list and `isWorkbookPlanIssueCode` guard, so agents
can branch on static handoff failures without inventing or string-matching
custom codes. Readback failures mirror that contract with
`WorkbookReadbackIssueCode`, `workbookReadbackIssueCodes`, and
`isWorkbookReadbackIssueCode`.
Formula readback expectation inputs must resolve through `refsUsed`, and
formula expectation text must be parseable.
Checks must start as `planned`; consumer model code cannot mark a check passed
or failed before runtime proof.
Low-level `addOp` commands must contain valid `WorkbookOp` values, must still
appear in `plan.ops`, and must match their declared `target` when the op exposes
a concrete address or range.
`verifyModel` applies the same planning and verification flow to every action
in a consumer-defined model, returning one JSON-safe model-level verdict.
`runWorkbookPlan(plan, adapter)` and
`runWorkbookAction(model, actionName, adapter, input)` add a transport-neutral
apply-and-prove loop on top of the same contracts. The adapter receives the full
plan, optionally previews materialized ops, applies it through whatever runtime
the consumer owns, and optionally returns semantic readbacks for the expectation
targets. `@bilig/workbook` compares those readbacks against `valueEquals` and
`formulaEquals` checks, plus the multi-cell `valuesEqual` and `formulasEqual`
checks, and returns a boring `WorkbookRunResult`. If static
verification fails, the apply adapter is not called. If preview is available,
successful run results include an `applied` summary with the materialized op
count and ops. If a readback expectation is missing, duplicated, or mismatched,
the run fails with deterministic codes such as `readback_missing`,
`duplicate_readback`, `value_mismatch`, `values_mismatch`, `formula_mismatch`,
or `formulas_mismatch`, plus structured `path`, `target`, `check`, `expected`,
and `actual` fields where available. Malformed adapter preview, apply, or
readback output fails as `invalid_runtime_result`; adapter output is not trusted
just because TypeScript says it has the right shape.
If apply succeeded and later proof fails, the failed run result preserves the
adapter undo ref when one was returned.
Formula readbacks are exact and should use the normalized no-leading-`=` form
produced by `formula.source`.
[`examples/workbook-agent-model`](../examples/workbook-agent-model) is the
small runnable proof for this flow. It defines a consumer-owned table model,
prints model/plan/runtime-requirement descriptions, executes through
`createWorkbookRunAdapter`, and returns a JSON-safe proof result. The example
depends on `@bilig/core` as a runtime choice; `@bilig/workbook` itself remains
runtime-free.
Run errors use the stable `WorkbookRunErrorCode` union. Agents and adapters can
inspect the frozen `workbookRunErrorCodes` list or call
`isWorkbookRunErrorCode(value)` before branching on a code. Runtime adapters
should use `apply_failed` for apply exceptions and `runtime_rejected` for
intentional runtime refusal with a specific message instead of inventing
model-specific public error codes.
`adapter.apply` only applies the plan and may return an undo ref; it cannot
drop, replace, or prove checks.
Adapters can also expose `verifyChecks(checks, plan)` for generic proof of
non-readback checks such as existence checks, formula-error checks, and
consumer-defined invariants. `verifyChecks` returns the same checks in the same
order and may only change `status`; malformed output fails with
`invalid_check_verification`, thrown verifier errors fail with
`check_verification_failed`, and failed checks become `check_failed` run errors.
If a check remains `planned` after readback and adapter verification,
`runWorkbookPlan` returns `failed` with `check_not_verified`; `status: "done"`
does not hide unproven checks.
`@bilig/core` provides `createWorkbookRunAdapter(engine)` for the canonical
engine handoff. It materializes generic `plan.commands` into engine operations,
including range and table-column writes, applies additional `plan.ops` that are
not already represented by materialized commands, previews exact materialized
ops before apply, reads single-cell and multi-cell readback targets, and verifies
generic `exists` and `noFormulaErrors` checks. When the engine
captures an undo transaction, the adapter returns a portable `undo.ops` ref
using the same workbook operation language. Consumer-defined business meaning
stays in the model; the core adapter only proves generic workbook facts.

## Core engine surface

The canonical engine surface includes:

- `createWorkbookRunAdapter`
- `createSheet`
- `deleteSheet`
- `setCellValue`
- `setCellFormula`
- `setCellFormat`
- `clearCell`
- `setRangeValues`
- `setRangeFormulas`
- `clearRange`
- `fillRange`
- `copyRange`
- `pasteRange`
- `setSelection`
- `undo`
- `redo`
- `getCell`
- `getDependencies`
- `getDependents`
- `explainCell`
- `exportSnapshot`
- `importSnapshot`
- `exportReplicaSnapshot`
- `importReplicaSnapshot`
- `applyRemoteBatch`
- `subscribe`
- `subscribeBatches`
- `connectSyncClient`
- `disconnectSyncClient`
- `getSyncState`

## WorkPaper surface

`@bilig/headless` exposes `WorkPaper`, a HyperFormula-style headless workbook API on top
of `@bilig/core`:

- `WorkPaper.buildEmpty`
- `WorkPaper.buildFromArray`
- `WorkPaper.buildFromSheets`
- `WorkPaper.buildFromSnapshot`
- workbook reads for cells, ranges, sheets, and named expressions
- display-value and formula-diagnostic reads for user-facing errors and
  structured financial validation details
- workbook mutations for cells, rows, columns, sheets, clipboard, and history
- `batch`, `suspendEvaluation`, `resumeEvaluation`, `undo`, and `redo`
- formula helpers such as address parsing, normalization, validation, and scratch evaluation
- static language and custom-function registration
- HyperFormula-style positional events through `on`, `once`, and `off`
- additive structured events through `onDetailed`, `onceDetailed`, and `offDetailed`
- stable internal adapter getters: `graph`, `rangeMapping`, `arrayMapping`,
  `sheetMapping`, `addressMapping`, `dependencyGraph`, `evaluator`,
  `columnSearch`, and `lazilyTransformingAstService`

`WorkPaper` is the canonical top-level contract.

## Excel Import Surface

`@bilig/headless/xlsx` exposes the CSV/XLSX boundary for WorkPaper consumers:

- `importXlsx(bytes, fileName)`
- `importCsv(text, fileName, options?)`
- `importWorkbookFile(bytes, fileName, contentType, options?)`
- `exportXlsx(snapshot)`
- `CSV_CONTENT_TYPE`
- `XLSX_CONTENT_TYPE`

CSV import auto-detects comma, semicolon, and tab delimiters. For locale-specific
accounting exports, pass `{ delimiter: ";", decimalSeparator: "," }`.

```sh
pnpm add @bilig/headless
```

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

const imported = importXlsx(new Uint8Array(readFileSync('model.xlsx')), 'model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  useColumnIndex: true,
})

const firstSheetName = imported.snapshot.sheets[0]?.name
const firstSheet = firstSheetName === undefined ? undefined : workbook.getSheetId(firstSheetName)
if (firstSheet === undefined) throw new Error('Workbook has no sheets')

workbook.setCellContents({ sheet: firstSheet, row: 1, col: 1 }, 150_000)
writeFileSync('model-edited.xlsx', exportXlsx(workbook.exportSnapshot()))
workbook.dispose()
```

Use `importXlsx(...).snapshot` with `WorkPaper.buildFromSnapshot()` when a
consumer needs Excel workbook metadata such as defined names, tables, and
translated structured references. `WorkPaper.buildFromSheets()` remains a
metadata-free array/sheet constructor. Use `workbook.exportSnapshot()` with
`exportXlsx()` to write an edited WorkPaper back to XLSX.

### Core types added in the current tranche

- `CellRangeRef`
- `SelectionRange`
- `SelectionEditMode`
- `SyncState`

## Binary protocol

`@bilig/binary-protocol` exposes:

- `PROTOCOL_VERSION`
- `encodeFrame(frame): Uint8Array`
- `decodeFrame(bytes): ProtocolFrame`

The current frame families are:

- `hello`
- `appendBatch`
- `ack`
- `snapshotChunk`
- `cursorWatermark`
- `heartbeat`
- `error`

The protocol surface is already real for sync and snapshot traffic. The remaining architecture gap is not “do we have binary frames?” but “does the authoritative workbook mutation language fully match what the local engine can represent?”

## Worker transport

`@bilig/worker-transport` exposes:

- `createWorkerEngineHost(engine, port)`
- `createWorkerEngineClient({ port })`

The first tranche already supports:

- method invocation
- engine events
- outbound batch subscriptions

## Target state

- the core API stays source-compatible while the worker-first runtime and durable backend land under it
- the binary protocol becomes the canonical wire format for browser sync, backend relay, and agent APIs
- the agent API moves from JSON payload bodies to typed binary request/response/event frames

## Exit gate

- every API documented here exists in code
- typecheck passes across all packages that import these APIs
- direct engine tests cover range mutation, history, selection state, and sync-state behavior
- the docs no longer claim a stable interface that is missing in the repo

## Agent API

`@bilig/agent-api` exposes:

- `AgentRequest`
- `AgentResponse`
- `AgentEvent`
- `encodeAgentFrame`
- `decodeAgentFrame`
- `encodeStdioMessage`
- `decodeStdioMessages`

Today the package defines typed TypeScript request/response/event unions, but the transport payload is still serialized with `JSON.stringify(...)` inside the binary frame envelope. The target state is a true typed binary schema shared by stdio and remote network usage.
