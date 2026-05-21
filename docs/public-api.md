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
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `verifyPlan`
- `verifyModel`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` through the model workbook context and as top-level helpers
- `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, and `check.custom` through the model workbook context and as top-level helpers
- `WorkbookModel`, `WorkbookAction`, `WorkbookActionInput`, `WorkbookAddOpOptions`, `WorkbookActionPlanResult`, `WorkbookModelDescription`, `WorkbookRefDescription`, `WorkbookActionPlanDescription`, `WorkbookActionPlanResultDescription`, `WorkbookPlanVerification`, `WorkbookPlanIssue`, `WorkbookModelVerification`, `WorkbookModelActionVerification`, `WorkbookModelVerificationOptions`, `WorkbookCheckExpectation`, `WorkbookCheckExpectationDescription`, `WorkbookCustomCheckOptions`, `WorkbookReadbackCheckOptions`, `WorkbookRawFormulaOptions`, `WorkbookRunResult`, and `WorkbookCheckResult`
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
does not provide schemas or validators for consumer meaning; actions keep that
validation generic and local. `verifyModel(model, { inputs })` supplies
per-action inputs for whole-model verification.

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

Action plans also expose `refsUsed`, a flat deduped list of workbook refs found
inside the consumer-defined `refs` object. This keeps custom models generic
while still letting agents inspect what the model resolved.
The same generic refs are available outside model callbacks through top-level
`findTable`, `findColumn`, `findRange`, `findName`, and `findRows` helpers.
`findRows` refs include their predicate value in the stable id and label, so
distinct consumer-defined row predicates remain distinct during agent
inspection and dedupe.
The same planned checks are available outside model callbacks through top-level
`check.exists(ref)`, `check.noFormulaErrors(ref)`,
`check.valueEquals(ref, value)`, `check.formulaEquals(ref, formula)`, and
`check.custom({ kind, message, target, refs })` helpers. Custom checks let
consumers carry their own invariants without adding hardcoded business models to
the package. `target` names the main ref, and `refs` names supporting refs so
agents can describe and verify the full invariant contract.
Readback checks add machine-readable expectations to the same generic check
channel: `valueEquals` stores the expected literal value, and `formulaEquals`
stores normalized formula text plus explicit formula input refs. Runtime code
can evaluate those expectations after applying the plan, while agents can
inspect the proof target without relying on visual spreadsheet state.

`describeModel` returns a JSON-safe model manifest with the model name, sorted
action names, and whether model-level checks exist. It does not run `find`,
checks, or actions.
For agent logs, approvals, tests, and runtime handoff, `describeRef` and
`describePlan` produce JSON-safe descriptions of refs and action plans. The
descriptions preserve generic action input and workbook intent while removing
consumer-private `refs` object shape and helper methods.
`describePlanResult` applies the same description layer to either planned or
failed action planning results.

`verifyPlan` gives agents a runtime-free consistency check before handoff. It
flags invalid action input, unresolved command targets, unresolved formula
inputs, duplicate resolved refs, unparsable formulas, and missing concrete ops
for write, clear, and number-format commands whose target is already known as a
single cell. Custom check targets and supporting refs must also resolve through
`refsUsed`.
Formula readback expectation inputs must resolve through `refsUsed`, and
formula expectation text must be parseable.
Low-level `addOp` commands must contain valid `WorkbookOp` values, must still
appear in `plan.ops`, and must match their declared `target` when the op exposes
a concrete address or range.
`verifyModel` applies the same planning and verification flow to every action
in a consumer-defined model, returning one JSON-safe model-level verdict.

## Core engine surface

The canonical engine surface includes:

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
