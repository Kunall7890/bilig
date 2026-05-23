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
- `workbookRefKinds`
- `isWorkbookRefKind`
- `isWorkbookRef`
- `workbookRowOperators`
- `isWorkbookRowOperator`
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRunResult`
- `describeRuntimeRequirements`
- `verifyPlan`
- `verifyModel`
- `workbookActionInputDescriptionKinds`
- `isWorkbookActionInputDescriptionKind`
- `isWorkbookActionInputDescription`
- `isWorkbookActionInput`
- `builtInWorkbookCheckKinds`
- `isBuiltInWorkbookCheckKind`
- `runWorkbookPlan`
- `runWorkbookAction`
- `verifyWorkbookReadbacks`
- `normalizeWorkbookActionInputDescription`
- `workbookRunErrorCodes`
- `isWorkbookRunErrorCode`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` through the model workbook context and as top-level helpers
- `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, and `check.custom` through the model workbook context and as top-level helpers
- `WorkbookModel`, `WorkbookAction`, `WorkbookActionConfig`, `WorkbookActionDefinition`, `WorkbookActionContext`, `WorkbookCheckContext`, `WorkbookFindWorkbook`, `WorkbookCheckWorkbook`, `WorkbookActionWorkbook`, `WorkbookModelWorkbook`, `WorkbookFindNamespace`, `WorkbookRef`, `WorkbookRefKind`, `WorkbookRangeRef`, `WorkbookNameRef`, `WorkbookTableRef`, `WorkbookColumnRef`, `WorkbookRowsRef`, `WorkbookRowOperator`, `WorkbookActionInput`, `WorkbookActionInputDescription`, `WorkbookActionInputDescriptionKind`, `WorkbookActionInspection`, `WorkbookAddOpOptions`, `WorkbookActionPlanResult`, `WorkbookModelDescription`, `WorkbookRefDescription`, `WorkbookActionPlanDescription`, `WorkbookActionPlanResultDescription`, `WorkbookRunResultDescription`, `WorkbookUndoRefDescription`, `WorkbookRuntimeRequirements`, `WorkbookRuntimeRequirement`, `WorkbookRuntimeCapability`, `WorkbookPlanVerification`, `WorkbookPlanIssue`, `WorkbookModelVerification`, `WorkbookModelActionVerification`, `WorkbookModelVerificationOptions`, `WorkbookRunAdapter`, `WorkbookRunApplyResult`, `WorkbookRunReadback`, `WorkbookReadbackVerification`, `WorkbookReadbackIssue`, `WorkbookReadbackIssueCode`, `WorkbookCheckExpectation`, `WorkbookCheckExpectationDescription`, `WorkbookBuiltInCheckKind`, `WorkbookCustomCheckOptions`, `WorkbookReadbackCheckOptions`, `WorkbookRawFormulaOptions`, `WorkbookRunResult`, `WorkbookRunError`, `WorkbookRunErrorCode`, and `WorkbookCheckResult`
- the existing low-level operation language: `WorkbookOp`, `WorkbookTxn`, `EngineOp`, and `EngineOpBatch`

The package builds portable workbook intent and concrete low-level ops when the
target is already known. Formula helpers use `@bilig/formula` for parsing and
normalization. Actual calculation and authoritative execution stay in
`@bilig/core` and `apps/bilig`.

`defineModel` returns frozen, normalized model metadata. Model and action names
must be non-empty and already trimmed, while descriptions and input metadata are
trimmed into model-owned frozen copies so the manifest an agent inspected cannot
be mutated later by the caller. The original consumer config remains
caller-owned data; `defineModel` does not freeze or rewrite it.
Action-object manifests only read own `run`, `description`, and `input`
properties. Prototype-inherited metadata is ignored, and an inherited `run`
function is rejected, so agent-visible manifests stay plain and explicit.

Model actions can accept plain JSON-safe input through
`planWorkbookAction(model, actionName, input)` and
`buildWorkbookActionPlan(model, actionName, input)`. The input is cloned,
canonicalized with stable object-key order, recorded on the plan, and exposed to
the action context as `input`. Supported values are strings, finite numbers,
booleans, `null`, arrays without holes, and plain objects. `@bilig/workbook`
does not provide schemas or validators for consumer meaning; actions keep that
validation generic and local. `verifyModel(model, { inputs })` supplies
per-action inputs for whole-model verification. The frozen
`workbookActionInputDescriptionKinds` list plus
`isWorkbookActionInputDescriptionKind(value)`,
`isWorkbookActionInputDescription(value)`, and `isWorkbookActionInput(value)` let
agents validate metadata and JSON-safe tool payloads without importing a schema
framework or string-matching ad hoc kinds.

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
`items`. `normalizeWorkbookActionInputDescription` trims text, rejects malformed
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

Formula helpers keep referenced workbook inputs separate from formula text.
Planned `writeFormula` commands expose those inputs directly, which lets agents
inspect dependencies without relying on human UI coordinates or reverse-parsing
placeholder formula names.
For formulas outside the small helper set, `formula.raw(source, { inputs })`
keeps arbitrary formula text generic while preserving explicit workbook
dependencies for inspection and verification. These are declared dependencies,
not parser-discovered proof that every formula reference has a matching model
ref.
Formula operands intentionally reject bare strings. Consumers use `formula.raw`
for formula source and `formula.text` for spreadsheet string literals, keeping
agent-authored formulas explicit instead of overloading a string as either code,
a label, a named range, or user text.

Action plans also expose `refsUsed`, a flat deduped list of workbook refs found
inside the consumer-defined `refs` object. This keeps custom models generic
while still letting agents inspect what the model resolved.
The same generic refs are available outside model callbacks through top-level
`findTable`, `findColumn`, `findRange`, `findName`, and `findRows` helpers, or
through the frozen `find` namespace with short aliases such as
`find.table(...)`, `find.range(...)`, and `find.rows(...)`.
The frozen `workbookRefKinds` and `workbookRowOperators` lists, plus
`isWorkbookRefKind`, `isWorkbookRef`, and `isWorkbookRowOperator`, expose the
same selector contract as data so agent tools can validate refs and row
predicates without copying string unions or pulling in a schema framework.
These selector helpers trim text, canonicalize cell addresses, and reject empty
or malformed selectors before the runtime handoff. That keeps bad agent intent
out of the plan instead of letting an invalid address, blank column, invalid row
operator, or non-finite predicate value fail later inside an engine adapter.
`findRows` refs include their predicate value in the stable id, so distinct
consumer-defined row predicates remain distinct during agent inspection and
dedupe. Labels stay simple and readable for logs.
Refs are frozen data objects. Helper methods such as `table.column()` and
`rows.column()` remain available for ergonomics, but they are non-enumerable so
object-key inspection and JSON descriptions stay data-first.
For table-backed row selectors, `rows.column("Amount")` targets that column only
inside the matching rows. Core/app runtime adapters can resolve those generic
refs into exact cells for writes, formats, clears, checks, and row-wise formula
input alignment without adding hardcoded workbook models.
The same planned checks are available outside model callbacks through top-level
`check.exists(ref)`, `check.noFormulaErrors(ref)`,
`check.valueEquals(ref, value)`, `check.formulaEquals(ref, formula)`, and
`check.custom({ kind, message, target, refs })` helpers. Custom checks let
consumers carry their own invariants without adding hardcoded business models to
the package. `target` names the main ref, and `refs` names supporting refs so
agents can describe and verify the full invariant contract.
Custom check kinds cannot reuse built-in names. Tool builders can inspect the
frozen `builtInWorkbookCheckKinds` list or call `isBuiltInWorkbookCheckKind`
before accepting consumer-defined check metadata.
Readback checks add machine-readable expectations to the same generic check
channel: `valueEquals` stores the expected literal value, and `formulaEquals`
stores normalized formula text plus explicit formula input refs. Runtime code
can evaluate those expectations after applying the plan, while agents can
inspect the proof target without relying on visual spreadsheet state.

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
Plans are frozen handoff objects: action input, refs used, commands, concrete
ops, changed summaries, and checks cannot be rewritten after planning. That
lets an agent inspect a plan once and pass the same intent to an adapter without
caller-side metadata drift.
`describePlanResult` applies the same description layer to either planned or
failed action planning results.
`describeRunResult` applies the same JSON-safe description layer after
execution, preserving `done`/`failed` status, changed summaries, checks, errors,
and undo ops while removing ref helper functions from the public result.
`describeRuntimeRequirements(plan)` gives agents a JSON-safe adapter checklist
for the same plan: which generic commands must be applied, which readbacks are
needed, and which checks need proof. It stays generic, with capabilities such
as `writeFormula`, `writeValue`, `format`, `clear`, `applyOp`, `read`, and
`verifyCheck`. Command-derived concrete single-cell ops are not repeated as
extra `applyOp` requirements, while explicit or manually assembled ops still
appear as `applyOp`. It does not import the engine.

`verifyPlan` gives agents a runtime-free consistency check before handoff. It
flags invalid action input, unresolved command targets, unresolved formula
inputs, duplicate resolved refs, unparsable formulas, and missing concrete ops
for write, clear, and number-format commands whose target is already known as a
single cell. Custom check targets and supporting refs must also resolve through
`refsUsed`.
Formula readback expectation inputs must resolve through `refsUsed`, and
formula expectation text must be parseable.
Checks must start as `planned`; consumer model code cannot mark a check passed
or failed before runtime proof.
Low-level `addOp` commands must contain valid `WorkbookOp` values, must still
appear in `plan.ops`, and must match their declared `target` when the op exposes
a concrete address or range.
`verifyModel` applies the same planning and verification flow to every action
in a consumer-defined model, returning one JSON-safe model-level verdict.
Successfully planned actions include their runtime requirements in the same
result, so an agent can inspect the planned intent, static proof, and adapter
handoff checklist without stitching multiple API calls together.
`runWorkbookPlan(plan, adapter)` and
`runWorkbookAction(model, actionName, adapter, input)` add a transport-neutral
apply-and-prove loop on top of the same contracts. The adapter receives the full
plan, applies it through whatever runtime the consumer owns, and optionally
returns semantic readbacks for the expectation targets. `@bilig/workbook`
compares those readbacks against `valueEquals` and `formulaEquals` checks and
returns a boring `WorkbookRunResult`. If static verification fails, the apply
adapter is not called. If a readback expectation is missing or mismatched, the
run fails with deterministic codes such as `readback_missing`,
`value_mismatch`, or `formula_mismatch`. Runtime readbacks must match the
requested target set exactly; surplus readbacks fail with `readback_unexpected`.
Formula readbacks are exact and should use the normalized no-leading-`=` form
produced by `formula.source`.
Run errors use the stable `WorkbookRunErrorCode` union. Agents and adapters can
inspect the frozen `workbookRunErrorCodes` list or call
`isWorkbookRunErrorCode(value)` before branching on a code. Runtime adapters
should use `apply_failed` for apply exceptions and `runtime_rejected` for
intentional runtime refusal with a specific message instead of inventing
model-specific public error codes.
`adapter.apply` only applies the plan and may return an undo ref; it cannot
drop, replace, or prove checks. Returning `status: "applied"` with non-empty
`errors` is rejected as `runtime_rejected`.
Adapters can also expose `verifyChecks(checks, plan)` for generic proof of
non-readback checks such as existence checks, formula-error checks, and
consumer-defined invariants. `verifyChecks` returns the same checks in the same
order and may only change `status` or add JSON-safe `proof`. Malformed output,
contract changes, invalid proof, and unsupported verifier mutations fail with
`invalid_check_verification`; thrown verifier errors fail with
`check_verification_failed`, and failed checks become `check_failed` run errors.
Accepted verifier output is sanitized before it reaches `WorkbookRunResult`. If
a check remains `planned` after readback and adapter verification,
`runWorkbookPlan` returns `failed` with `check_not_verified`; `status: "done"`
does not hide unproven checks.
`@bilig/core` provides `createWorkbookRunAdapter(engine)` for the canonical
engine handoff. It materializes generic `plan.commands` into engine operations,
including range and table-column writes, falls back to explicit `plan.ops` for
low-level plans, reads single-cell `valueEquals` and `formulaEquals` targets,
and verifies generic `exists` and `noFormulaErrors` checks. When the engine
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
