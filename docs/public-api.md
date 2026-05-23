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

The stable contract has four layers.

### Agent model contract

Consumers define models. Bilig does not ship hardcoded revenue, quote, forecast,
prepaid, or other business models in this package.

- `defineModel` freezes model metadata and action manifests.
- `find` returns generic workbook refs through `findTable`, `findColumn`,
  `findRange`, `findName`, and `findRows`.
- `checks` declares generic facts the runtime must prove.
- `actions` build portable workbook intent with `writeFormula`, `writeValue`,
  `format`, `clear`, or a guarded low-level op.
- `planWorkbookAction` returns either a frozen plan or a structured failure.
- `verifyPlan` proves static consistency without importing or starting an
  engine.
- `describeModel`, `describePlan`, and `describePlanResult` return JSON-safe
  objects for logs, approvals, tools, and tests.
- `toPlanData` returns JSON-safe plan data for runtime handoff.
- `checkPlanData` returns structured path-based issues for transported plan
  payloads before hydration.
- `hydratePlanData` restores frozen refs and helper methods from transported
  plan data without reconstructing the consumer's private `refs` shape.
- `verifyPlanData` proves a JSON-transported plan description after it has lost
  local helper methods and consumer-private `refs` object shape.

The model contract is intentionally data-first. Refs are frozen. Helper methods
such as `table.column("Amount")` and `rows.column("Amount")` are non-enumerable.
Action input is plain JSON, cloned into the plan, and described with small
metadata rather than a schema dependency.
Action manifests are frozen null-prototype maps, and planning only runs own
manifest actions. Generic action names such as `toString`, `constructor`, or
`__proto__` can be explicit own actions, but inherited prototype fields are not
part of the executable surface.
For transport, `toWorkbookRefData` turns any ref into plain data,
`isWorkbookRefData` validates that data, `collectWorkbookRefData` discovers refs
without requiring hidden methods, and `hydrateWorkbookRef` /
`hydrateWorkbookRefs` restores the ergonomic helpers when local code needs them
again.
Ref discovery and ref hydration only inspect enumerable own data properties.
Accessor properties are ignored instead of invoked, so consumer-defined getters
cannot run during planning, verification, logging, or transport hydration.
Array entries follow the same data-only rule, and ref cloning copies known ref
fields rather than spreading extra enumerable properties.
Selector creation is data-only too. `findTable`, `findColumn`, `findRange`, and
`findRows` reject accessor-backed option fields, row predicates, and header
entries before any hidden getter can run.
For full action handoff, `toPlanData(plan)` returns executable JSON-safe plan
data. Runtimes can call `hydratePlanData(data)`, or pass that data directly to
`describeRuntimeRequirements(data)` and `runWorkbookPlan(data, adapter)`.
Hydrated transported plans expose `refs: { refsUsed }`, keeping execution
generic instead of depending on the consumer's model-shaped `refs` object.
Agents that receive untrusted JSON can call `checkPlanData(data)` first to get
boring `{ status, issues }` diagnostics such as `commands[0]` or `refsUsed[2]`
instead of catching a generic hydration failure. Transported plan arrays are
also data-only: holes, non-enumerable entries, or accessor-backed entries are
rejected without invoking getters before hydration or execution.

### Runtime adapter contract

Runtimes execute plans. `@bilig/workbook` only defines the handoff.

- `describeRuntimeRequirements(plan)` tells an adapter which commands must be
  applied, which targets must be read back, and which checks need runtime proof.
- `checkRuntimeRequirements(data)` returns structured path-based issues for
  transported runtime requirement payloads before adapter validation.
- `runWorkbookPlan(planOrData, adapter)` accepts either the in-memory plan or
  transported plan data, refuses to apply invalid plans, calls the adapter,
  verifies readback expectations, verifies generic checks, and returns a boring
  `WorkbookRunResult`.
- `checkRuntimeAdapter(planOrRequirements, adapter)` checks that the adapter has
  the required `apply`, `read`, and `verifyChecks` methods before mutation.
- `adapter.apply(plan)` owns runtime mutation and may return `previewOps`,
  `appliedOps`, JSON-safe `proof`, and an undo ref.
- `adapter.read(targets, plan)` returns semantic readbacks for checks such as
  `valueEquals` and `formulaEquals`.
- `adapter.verifyChecks(checks, plan)` may only change check `status` or add
  JSON-safe `proof`; it cannot rewrite the check contract.

Adapter-returned apply results, undo refs, apply errors, and check verifier
output are accepted from own payload fields only. Prototype-inherited fields do
not satisfy runtime proof. Adapter-returned op arrays and verifier proof must be
plain data properties; runtime evidence arrays must contain own enumerable data
entries. Holes, non-enumerable entries, and accessor-backed fields are rejected
before cloning or preview/apply comparison, including non-enumerable fields that
op guards would otherwise read directly.

When `previewOps` and `appliedOps` are both present, `runWorkbookPlan` reports
whether runtime apply matched preview. When they are missing, the result reports
an unverified apply fact. Agents that need fail-closed execution can call
`runWorkbookPlan(plan, adapter, { requireApplyProof: true })`.
Plans with no apply requirements skip mutation entirely: a readback-only or
check-only model can run with only `read` or `verifyChecks`, and the result has
no `apply` summary because nothing was supposed to mutate.
Failed results include `changed` as a concrete answer, not an implied absence:
`changed: []` means no runtime apply succeeded before the failure, while
post-apply proof failures keep the planned change summaries and adapter-provided
undo metadata. If an adapter returns a failed apply with `appliedOps: []` and no
undo metadata, `changed` stays empty because the runtime proved no ops were
applied.

Readback-backed checks attach proof to passed checks. A result can therefore
show the intended action, bound refs, planned commands and ops, adapter
requirements, preview ops, applied ops, preview/apply match, readback values or
formulas, check statuses, proof objects, undo metadata, and remaining unverified
facts without relying on rendered spreadsheet state.

### Feature command handoff

Runtimes may expose workbook extensions as commands, projection interceptors,
and UI contributions. The public package still stays data-only.

- `defineWorkbookFeaturePlugin` freezes extension metadata.
- `checkWorkbookFeaturePlugin(data)` validates consumer-provided feature
  manifests before registration, including nested command input-description and
  UI contribution metadata paths.
- `checkWorkbookCommandRequest(data)` validates transported command requests
  before dispatch and returns path issues such as `featureId`, `commandId`,
  `category`, `mode`, and nested JSON payload paths such as `input.rows[1]`.
- `normalizeWorkbookCommandRequest(data)` returns a frozen request after the
  same validation.
- `checkWorkbookCommandReceipt(data)` validates transported receipt evidence
  before an agent trusts runtime extension output, including nested `proof` and
  `metadata` payload paths.
- `normalizeWorkbookCommandReceipt(receipt)` and
  `workbookCommandReceiptOpsMatch(receipt)` give agents boring receipt proof
  after preview or apply. Receipt ops are normalized into frozen data and
  compared by canonical op content, so property insertion order cannot create a
  false mismatch and invalid op arrays cannot report a trusted match.
- Feature manifests, command requests, and command receipts are validated from
  own payload fields only; prototype-inherited fields cannot satisfy the public
  transport contract. Manifest and receipt arrays must contain own enumerable
  data entries; receipt ops, undo ops, changed ranges, and errors must be data
  properties. Holes, non-enumerable entries, and accessor-backed receipt proof
  are rejected before freezing or preview/apply comparison.
- Frozen vocabularies such as `workbookCommandCategories`,
  `workbookCommandExecutionModes`, `workbookCommandReceiptStatuses`,
  `workbookProjectionInterceptorPoints`, and `workbookUiContributionSlots` let
  tool builders expose exact supported values without a schema framework.

Feature command handoff does not move execution into `@bilig/workbook`. The
runtime owns command semantics. The package only normalizes the manifest,
validates transported requests, and describes receipt evidence.

### Escape hatches

The preferred path is the small model API. Escape hatches stay explicit:

- `findRange` is available when a consumer really has a concrete range.
- `formula.raw(source, { inputs })` is available for formulas outside the helper
  set while keeping dependencies inspectable. Add `labels: [{ name, ref }]`
  when raw formula text uses custom ref tokens.
- `workbook.addOp(op, { target?, message? })` carries the existing low-level
  workbook operation language through the same plan, description, and
  verification flow.

Escape hatches do not make the package domain-specific. They keep the public API
generic while still letting advanced runtimes use the lower-level workbook
language.
Low-level op guards accept only plain own-field payloads. Prototype-inherited
op fields, nested ranges, and batch clocks do not satisfy `isWorkbookOp` or
`isEngineOpBatch`. Accessor-backed required fields, nested fields, and op-array
entries are rejected from descriptors without invoking getters.

The generic runnable example lives in
[`examples/workbook-agent-model`](../examples/workbook-agent-model). Domain
examples such as revenue, forecast, and quote approval belong under consumer app
or WorkPaper examples, not inside `@bilig/workbook` itself.

Full export surface:

- `defineModel`
- `buildWorkbookActionPlan`
- `planWorkbookAction`
- `inspectModel`
- `collectWorkbookRefs`
- `collectWorkbookRefData`
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows`
- `find`
- `workbookRefKinds`
- `isWorkbookRefKind`
- `isWorkbookRef`
- `isWorkbookRefData`
- `toWorkbookRefData`
- `hydrateWorkbookRef`
- `hydrateWorkbookRefs`
- `workbookRowOperators`
- `workbookRowOperatorValueTypes`
- `isWorkbookRowOperator`
- `isWorkbookRowValueCompatible`
- `check`
- `describeModel`
- `describeRef`
- `describePlan`
- `describePlanResult`
- `describeRunResult`
- `describeRuntimeRequirements`
- `checkRuntimeRequirements`
- `checkRuntimeAdapter`
- `toPlanData`
- `isPlanData`
- `checkPlanData`
- `hydratePlanData`
- `verifyPlan`
- `verifyPlanData`
- `verifyModel`
- `checkInput`
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
- `workbookRuntimeRequirementKinds`
- `isWorkbookRuntimeRequirementKind`
- `workbookRuntimeCapabilities`
- `isWorkbookRuntimeCapability`
- `defineWorkbookFeaturePlugin`
- `checkWorkbookFeaturePlugin`
- `checkWorkbookCommandRequest`
- `normalizeWorkbookCommandRequest`
- `isWorkbookCommandRequest`
- `checkWorkbookCommandReceipt`
- `normalizeWorkbookCommandReceipt`
- `isWorkbookCommandReceipt`
- `workbookCommandReceiptOpsMatch`
- `workbookCommandCategories`
- `isWorkbookCommandCategory`
- `workbookCommandExecutionModes`
- `isWorkbookCommandExecutionMode`
- `workbookCommandReceiptStatuses`
- `isWorkbookCommandReceiptStatus`
- `workbookProjectionInterceptorPoints`
- `isWorkbookProjectionInterceptorPoint`
- `workbookUiContributionSlots`
- `isWorkbookUiContributionSlot`
- `formula`
- `workbook.addOp(op, { target?, message? })` inside model actions
- `findTable`, `findColumn`, `findRange`, `findName`, and `findRows` through the model workbook context and as top-level helpers
- `check.exists`, `check.noFormulaErrors`, `check.valueEquals`, `check.formulaEquals`, and `check.custom` through the model workbook context and as top-level helpers
- `WorkbookModel`, `WorkbookAction`, `WorkbookActionConfig`, `WorkbookActionDefinition`, `WorkbookActionContext`, `WorkbookCheckContext`, `WorkbookFindWorkbook`, `WorkbookCheckWorkbook`, `WorkbookActionWorkbook`, `WorkbookModelWorkbook`, `WorkbookFindNamespace`, `WorkbookRef`, `WorkbookRefData`, `WorkbookRefKind`, `WorkbookBaseRefData`, `WorkbookRangeRef`, `WorkbookRangeRefData`, `WorkbookNameRef`, `WorkbookNameRefData`, `WorkbookTableRef`, `WorkbookTableRefData`, `WorkbookColumnRef`, `WorkbookColumnRefData`, `WorkbookRowsRef`, `WorkbookRowsRefData`, `WorkbookRowOperator`, `WorkbookRowValueType`, `WorkbookActionInput`, `WorkbookActionInputDescription`, `WorkbookActionInputDescriptionKind`, `WorkbookActionInputIssueCode`, `WorkbookActionInputIssue`, `WorkbookActionInputCheckResult`, `WorkbookActionInspection`, `WorkbookAddOpOptions`, `WorkbookActionPlanResult`, `WorkbookModelDescription`, `WorkbookRefDescription`, `WorkbookActionPlanDescription`, `WorkbookPlanData`, `WorkbookPlanDataRefs`, `WorkbookPlanDataIssueCode`, `WorkbookPlanDataIssue`, `WorkbookPlanDataCheckResult`, `WorkbookExecutablePlan`, `WorkbookActionPlanResultDescription`, `WorkbookRunResultDescription`, `WorkbookUndoRefDescription`, `WorkbookRunApplySummaryDescription`, `WorkbookRunUnverifiedDescription`, `WorkbookRuntimeRequirements`, `WorkbookRuntimeRequirement`, `WorkbookRuntimeRequirementKind`, `WorkbookRuntimeRequirementsIssueCode`, `WorkbookRuntimeRequirementsIssue`, `WorkbookRuntimeRequirementsCheckResult`, `WorkbookRuntimeCapability`, `WorkbookRuntimeAdapterIssueCode`, `WorkbookRuntimeAdapterMethod`, `WorkbookRuntimeAdapterIssue`, `WorkbookRuntimeAdapterCheckResult`, `WorkbookRuntimeAdapterCandidate`, `WorkbookPlanVerification`, `WorkbookPlanIssue`, `WorkbookModelVerification`, `WorkbookModelActionVerification`, `WorkbookModelVerificationOptions`, `WorkbookRunAdapter`, `WorkbookRunApplyResult`, `WorkbookRunOptions`, `WorkbookRunApplySummary`, `WorkbookRunUnverified`, `WorkbookRunUnverifiedKind`, `WorkbookRunReadback`, `WorkbookReadbackVerification`, `WorkbookReadbackIssue`, `WorkbookReadbackIssueCode`, `WorkbookCheckExpectation`, `WorkbookCheckExpectationDescription`, `WorkbookBuiltInCheckKind`, `WorkbookCustomCheckOptions`, `WorkbookReadbackCheckOptions`, `WorkbookFormulaExpression`, `WorkbookFormulaLabel`, `WorkbookFormulaLabelDescription`, `WorkbookRawFormulaOptions`, `WorkbookRunResult`, `WorkbookRunError`, `WorkbookRunErrorCode`, and `WorkbookCheckResult`
- `WorkbookFeatureId`, `WorkbookCommandCategory`, `WorkbookCommandExecutionMode`, `WorkbookCommandReceiptStatus`, `WorkbookProjectionInterceptorPoint`, `WorkbookUiContributionSlot`, `WorkbookFeatureLifecycleContext`, `WorkbookCommandDescriptor`, `WorkbookCommandRequest`, `WorkbookCommandRequestIssueCode`, `WorkbookCommandRequestIssue`, `WorkbookCommandRequestCheckResult`, `WorkbookCommandReceipt`, `WorkbookCommandReceiptIssueCode`, `WorkbookCommandReceiptIssue`, `WorkbookCommandReceiptCheckResult`, `WorkbookFeaturePluginIssueCode`, `WorkbookFeaturePluginIssue`, `WorkbookFeaturePluginCheckResult`, `WorkbookCellDisplayProjection`, `WorkbookCellStyleProjection`, `WorkbookRangeChromeProjection`, `WorkbookRowVisibilityProjection`, `WorkbookCommandMetadataProjection`, `WorkbookProjectionContext`, `WorkbookProjectionInterceptorRegistration`, `WorkbookUiContribution`, `WorkbookFeatureRegistration`, and `WorkbookFeaturePlugin`
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
Accessor-backed model config, action-map entries, and action-object metadata are
rejected without invoking hidden getters.

Model actions can accept plain JSON-safe input through
`planWorkbookAction(model, actionName, input)` and
`buildWorkbookActionPlan(model, actionName, input)`. The input is cloned,
canonicalized with stable object-key order, recorded on the plan, and exposed to
the action context as `input`. Supported values are strings, finite numbers,
booleans, `null`, arrays without holes, and plain objects. `@bilig/workbook`
does not provide schemas for consumer meaning; actions keep domain validation
local. `checkInput(description, value)` only checks the package's small generic
input metadata and returns a plain `{ status, input, issues }` result.
Omitted input is valid unless the top-level description sets `required: true`;
required omissions return `missing_required_input` instead of pretending
`undefined` is a malformed JSON payload. `planWorkbookAction` runs that same
check before `find`, `checks`, or action code when an action declares input
metadata. Failed action input checks keep the stable input issue `path` and
`issueCode` on each run error, so agents can branch on structured diagnostics
without parsing messages. JSON-safety failures preserve nested paths like
`input.items[2].amount` instead of collapsing every issue to `input`.
Normalized payloads preserve consumer-owned JSON keys as own data properties,
including names like `__proto__` and `constructor`, so transported tool payloads
cannot mutate prototypes or disappear during canonicalization.
Action input payloads and input-description metadata must be enumerable own data
properties. Accessors are rejected without invocation, which keeps tool payload
validation inspectable and prevents hidden consumer code from running during
planning.
`verifyModel(model, { inputs })` supplies per-action inputs for
whole-model verification. The frozen
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
`items`; any description can mark itself `required`. `checkInput` allows
unknown object fields, because consumer actions own consumer-specific meaning.
`normalizeWorkbookActionInputDescription` trims text, rejects malformed metadata,
freezes the result, and keeps `@bilig/workbook` independent from `zod`, `effect`,
and model-specific validators.

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
Raw `WorkbookOp` and `EngineOpBatch` guards trust own fields only, including
nested cell ranges and batch clocks, so a runtime can reject prototype-shaped
payloads before hydration or execution. They also reject accessor-backed
required fields, nested fields, and op-array entries without running getters.

Formula helpers keep referenced workbook inputs and formula labels separate from
formula text. Planned `writeFormula` commands expose `inputs` plus `labels`,
where each label maps a token in the formula string to the workbook ref it
represents. Agents can inspect and verify that handoff without relying on human
UI coordinates, hidden JS helper calls, or reverse-parsing placeholder names.
For formulas outside the small helper set, `formula.raw(source, { inputs })`
keeps arbitrary formula text generic while preserving explicit workbook
dependencies for inspection and verification. Use
`formula.raw(source, { inputs, labels })` when the raw formula uses custom ref
tokens. These are declared dependencies and token mappings, not a
parser-discovered model.
Formula operands intentionally reject bare strings. Consumers use `formula.raw`
for formula source and `formula.text` for spreadsheet string literals, keeping
agent-authored formulas explicit instead of overloading a string as either code,
a label, a named range, or user text.

Action plans also freeze the consumer-defined `refs` container graph and expose
`refsUsed`, a flat deduped list of workbook refs found inside that object. This
keeps custom models generic while still letting agents inspect what the model
resolved.
The same generic refs are available outside model callbacks through top-level
`findTable`, `findColumn`, `findRange`, `findName`, and `findRows` helpers, or
through the frozen `find` namespace with short aliases such as
`find.table(...)`, `find.range(...)`, and `find.rows(...)`.
The frozen `workbookRefKinds`, `workbookRowOperators`, and
`workbookRowOperatorValueTypes` data, plus `isWorkbookRefKind`,
`isWorkbookRef`, `isWorkbookRowOperator`, and `isWorkbookRowValueCompatible`,
expose the same selector contract as data so agent tools can validate refs and
row predicates without copying string unions or pulling in a schema framework.
These selector helpers trim text, canonicalize cell addresses, and reject empty
or malformed selectors before the runtime handoff. That keeps bad agent intent
out of the plan instead of letting an invalid address, blank column, duplicate
header, invalid row operator, non-finite predicate value, or invalid
operator/value pair fail later inside an engine adapter.
Table header selectors are an all-of match, not a coordinate or position
contract. Headers are case-sensitive after trimming, stored in sorted order, and
deduped so agents get stable ids for the same generic table intent.
Row selectors keep the value contract simple: `eq` and `neq` accept any JSON
literal, `contains` and `startsWith` accept strings, and ordered comparisons
accept numbers or strings.
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
stores normalized formula text plus explicit formula input refs and labels.
Runtime code can evaluate those expectations after applying the plan, while
agents can inspect the proof target without relying on visual spreadsheet state.

Model callback phases expose only the API needed for that phase. `find` gets
find helpers only, `checks` gets find helpers and planned-check helpers, and
actions get find helpers, checks, and mutation-planning methods. Discovery and
proof declaration cannot accidentally write workbook intent before the action
phase.
Checks returned from a consumer `checks()` callback are also treated as a data
boundary. Returned arrays must contain own enumerable data entries, and returned
check fields must be own data properties; sparse arrays and accessor-backed
checks fail planning without invoking hidden getters.

`describeModel` returns a JSON-safe model manifest with the model name, optional
model description, sorted action names, per-action descriptions, optional input
descriptions, and whether model-level checks exist. It does not run `find`,
checks, or actions.
For agent logs, approvals, tests, and runtime handoff, `describeRef` and
`describePlan` produce JSON-safe descriptions of refs and action plans. The
descriptions preserve generic action input and workbook intent while removing
consumer-private `refs` object shape and helper methods.
Those descriptions are real transport data, not screenshots or UI state:
`isWorkbookRefData` validates ref payloads, `hydrateWorkbookRefs` restores
local helper methods after JSON transport, and `toPlanData` makes a full plan
JSON-safe for handoff. `isPlanData` validates that payload, `checkPlanData`
explains invalid payloads with stable paths, `hydratePlanData` restores frozen
refs and helpers, and `verifyPlanData` verifies transported plan data using only
`refsUsed`, commands, ops, changes, and checks. Invalid transported action input
and check proof keep nested JSON paths such as `input.rows[1]` and
`checks[0].proof.when`. Plan-data validation treats transport payloads as own
data only, so prototype-inherited fields cannot make an invalid payload look
valid.
Plans are frozen handoff objects: action input, refs, refs used, commands,
concrete ops, changed summaries, and checks cannot be rewritten after planning.
That lets an agent inspect a plan once and pass the same intent to an adapter
without caller-side metadata drift.
`describePlanResult` applies the same description layer to either planned or
failed action planning results.
`describeRunResult` applies the same JSON-safe description layer after
execution, preserving `done`/`failed` status, changed summaries, checks, errors,
and undo ops while removing ref helper functions from the public result.
`describeRuntimeRequirements(plan)` gives agents a JSON-safe adapter checklist
for the same plan or transported plan data: which generic commands must be applied, which readbacks are
needed, and which checks need proof. It stays generic, with capabilities such
as `writeFormula`, `writeValue`, `format`, `clear`, `applyOp`, `read`, and
`verifyCheck`. Command-derived concrete single-cell ops are not repeated as
extra `applyOp` requirements, while explicit or manually assembled ops still
appear as `applyOp`. It does not import the engine.
`checkRuntimeRequirements(data)` validates that checklist after JSON transport
and returns boring `{ status, issues }` diagnostics with paths such as
`requirements[0].capability` or `requirements[2].refs[0]`. That lets agents
reject malformed adapter handoff data before checking runtime methods or
starting mutation. Runtime requirement validation also ignores inherited fields;
the adapter checklist has to be present as explicit payload data. Requirement
arrays and nested ref arrays are data-only too: holes, non-enumerable entries,
or accessor-backed entries are rejected without invoking getters.

`verifyPlan` gives agents a runtime-free consistency check before handoff. It
flags invalid action input, unresolved command targets, unresolved formula
inputs, missing formula labels, formula labels that do not point at resolved
refs or do not appear in formula text, duplicate resolved refs, refs used that
are not discoverable from `refs`, unparsable formulas, and missing concrete ops
for write, clear, and number-format commands whose target is already known as a
single cell. Custom check targets and supporting refs must also resolve through
`refsUsed`. Formula readback expectation inputs and labels must resolve through
`refsUsed`, and formula expectation text must be parseable.
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
`checkRuntimeAdapter(planOrRequirements, adapter)` compares that checklist to an
adapter shape and returns a plain valid/invalid result with missing capability
issues. It accepts a live plan, transported plan data, or the output of
`describeRuntimeRequirements`.
`runWorkbookPlan(planOrData, adapter)` and
`runWorkbookAction(model, actionName, adapter, input)` add a transport-neutral
apply-and-prove loop on top of the same contracts. The adapter receives the full
plan, or a hydrated transported plan with `refs: { refsUsed }`. When the
requirements include an apply capability, the adapter applies it through
whatever runtime the consumer owns and may return `previewOps`, `appliedOps`,
apply proof, undo metadata, and semantic readbacks for the expectation targets.
When the plan only requires readbacks or generic check proof, `@bilig/workbook`
skips mutation and runs those proof steps directly. `@bilig/workbook` compares
preview ops to applied ops when both are present, compares readbacks against
`valueEquals` and `formulaEquals` checks, and returns a boring
`WorkbookRunResult`. Passed readback-backed checks include JSON-safe proof such as
`{ source: "readback", value }` or `{ source: "readback", formula }`, so the
result says what the runtime actually read back. If static verification fails,
the apply adapter is not called. If the adapter is missing a required method for
the plan, the apply adapter is not called and the run fails with
`adapter_missing_capability`. If preview/apply ops mismatch, the run fails with
`apply_mismatch`. If `requireApplyProof` is true and the adapter omits preview
or applied ops, the run fails with `apply_not_verified`. If a readback
expectation is missing or mismatched after a reader runs, the run fails with
deterministic codes such as `readback_missing`, `value_mismatch`, or
`formula_mismatch`.
Failures after an adapter reports `status: "applied"` preserve `changed` and
`undo`, so an agent can still inspect what was applied and how to reverse it
when a later proof step rejects the run.
Failed apply results only preserve `changed` when the adapter reports applied
ops or undo metadata; an explicit empty `appliedOps` array with no undo keeps
`changed: []`.
Runtime readbacks must match the requested target set exactly; surplus
readbacks fail with `readback_unexpected`. Formula readbacks are exact and
should use the normalized no-leading-`=` form produced by `formula.source`.
Run errors use the stable `WorkbookRunErrorCode` union. Agents and adapters can
inspect the frozen `workbookRunErrorCodes` list or call
`isWorkbookRunErrorCode(value)` before branching on a code. Runtime adapters
should use `apply_failed` for apply exceptions and `runtime_rejected` for
intentional runtime refusal with a specific message instead of inventing
model-specific public error codes. Errors may also include a stable `path` and
`issueCode` when a generic checker can identify the exact rejected input or
adapter issue.
`adapter.apply` only applies the plan and may return apply proof plus an undo
ref; it cannot drop, replace, or prove checks. Returning `status: "applied"`
with non-empty `errors` is rejected as `runtime_rejected`. If no apply proof is
required and the adapter omits preview or applied ops, the run can still finish
but reports an `unverified` apply fact instead of pretending preview/apply match
was proven.
Apply result fields, nested runtime errors, and undo metadata are sanitized from
own fields before they reach `WorkbookRunResult`. Accessor-backed preview ops,
applied ops, undo ops, runtime errors, and verifier proof are rejected without
invoking getters.
If an adapter returns `status: "failed"` with `appliedOps` or `undo`, the failed
run preserves the planned change summaries and undo ref because the runtime is
signaling that mutation evidence exists even though the apply step rejected the
overall run.
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
Verifier output is also own-field only: inherited `status`, `kind`, `message`,
target, refs, expectation, or proof data is ignored and cannot prove a check.
`@bilig/core` provides `createWorkbookRunAdapter(engine)` for the canonical
engine handoff. It materializes generic `plan.commands` into engine operations,
including range and table-column writes, falls back to explicit `plan.ops` for
low-level plans, reads single-cell `valueEquals` and `formulaEquals` targets,
and verifies generic `exists` and `noFormulaErrors` checks. When the engine
applies a plan, the adapter returns matching `previewOps` and `appliedOps`, plus
JSON-safe apply proof. When the engine captures an undo transaction, the adapter
returns a portable `undo.ops` ref using the same workbook operation language.
Consumer-defined business meaning stays in the model; the core adapter only
proves generic workbook facts.

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
