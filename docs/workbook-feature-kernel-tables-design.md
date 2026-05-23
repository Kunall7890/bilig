# Workbook Feature Kernel With Tables Product Slice

## Status

Design target for the production implementation of the workbook feature kernel and the first tables-backed product slice.

## Problem

Workbook features in bilig already have strong underlying engine support, but the ownership boundaries are uneven. Table metadata, command behavior, UI actions, projection behavior, agent preview summaries, and undo/proof handling are spread across engine, agent API, app shell, and toolbar code. That makes each new workbook feature harder than it should be, and it increases the risk that UI-only behavior diverges from authoritative workbook state.

The implementation must introduce a reusable feature architecture without replacing bilig strengths:

- formula evaluation stays in the existing formula engine
- rendering stays on the TypeGPU renderer path
- collaboration and local mutation replay stay on the existing mutation journal and sync paths
- workbook actions keep the existing preview/apply/verify/undo proof model
- table storage stays in the existing workbook metadata schema for this slice

## Goals

- Add a typed workbook feature plugin contract that lets feature code own commands, projection interceptors, UI contributions, and lifecycle.
- Add a command service that treats user actions, engine operations, and atomic mutations as explicit concepts while still executing through current engine op APIs.
- Add projection interceptors so features can project display/style/chrome/visibility metadata without mutating cells.
- Add a typed facade for product and agent code that can target workbook concepts such as tables, columns, selections, and commands.
- Add lightweight web UI slots so feature UI can be mounted without expanding the app shell or toolbar files.
- Prove the architecture with a complete tables slice: create from selection, rename headers, resize, delete, inspect, project chrome, preview/apply/verify, undo, and agent semantic targeting.

## Non-Goals

- Do not replace the engine, formula runtime, renderer, snapshots, import/export metadata format, Zero sync, or mutation journal.
- Do not create a separate feature package in the first slice. Use current package boundaries and extract packages after at least two features share the kernel cleanly.
- Do not change the public agent bundle wire format for existing table commands.
- Do not add visual table chrome unless it is derived from authoritative table metadata.

## Architecture

### Workbook Feature Contract

`@bilig/workbook` owns the stable feature contract because it is the lowest shared workbook API layer.

Required public types:

- `WorkbookFeaturePlugin`
- `WorkbookFeatureLifecycleContext`
- `WorkbookFeatureRegistration`
- `WorkbookCommandDescriptor`
- `WorkbookCommandCategory`
- `WorkbookCommandRequest`
- `WorkbookCommandReceipt`
- `WorkbookCommandExecutionMode`
- `WorkbookProjectionInterceptorRegistration`
- `WorkbookProjectionInterceptorPoint`
- `WorkbookUiContribution`

The plugin contract has:

- `id`: stable string identifier, for example `tables`
- `version`: semantic version string owned by the feature
- `dependsOn`: optional list of feature ids that must be active first
- `commands`: command descriptors registered by the feature
- `projectionInterceptors`: pure ordered projection handlers
- `uiContributions`: serializable contribution descriptors consumed by app code
- `register(context)`, `activate(context)`, and `dispose(context)` lifecycle hooks

Lifecycle rules:

- Duplicate feature ids fail during registration.
- Missing dependencies fail before activation.
- Dependencies activate before dependents.
- `dispose` runs in reverse activation order.
- Lifecycle methods are idempotent from the registry caller's point of view.

### Command Service

`@bilig/core` owns runtime command execution because it can safely see the `SpreadsheetEngine`.

Command categories:

- `command`: high-level user or agent intent, such as create table from selection
- `operation`: domain operation that expands to engine ops, such as resize table
- `mutation`: already-atomic engine op or op batch

Every execution returns a `WorkbookCommandReceipt`:

- `status`: `previewed`, `applied`, `rejected`, or `noop`
- `featureId`
- `commandId`
- `category`
- `previewOps`
- `appliedOps`
- `undo`
- `changedRanges`
- `proof`
- `message`
- `metadata`

Execution requirements:

- Preview and apply must use the same command descriptor.
- Apply must capture undo information through current engine mutation/undo APIs.
- If preview ops and applied ops differ, the receipt must say so explicitly and tests must cover it.
- Rejected commands return structured errors and do not mutate the workbook.
- Existing engine protections and metadata validation remain authoritative.

### Projection Interceptors

`@bilig/core` owns the projection service; feature plugins contribute interceptors.

Projection points:

- `cellDisplay`: display value or display metadata for a cell
- `cellStyle`: projected style hints derived from metadata
- `rangeChrome`: range-level visual chrome such as table borders or headers
- `rowVisibility`: visibility state derived from feature metadata
- `beforeCommand`: command gate or command metadata enrichment
- `commandMetadata`: labels, affected ranges, semantic refs, and proof hints

Rules:

- Interceptors are pure. They do not mutate workbook state.
- Interceptors are ordered by priority, then registration order.
- The service returns both projected output and provenance so UI/tests can prove it came from authoritative metadata.
- Table chrome and header projection must be computed from `WorkbookTableSnapshot` records, never from UI-only state.

### Facade

`@bilig/core` exposes a typed `WorkbookFacade` on top of the command and projection services.

Required v1 methods:

- `table(name).snapshot()`
- `table(name).column(columnName)`
- `selection(range).createTable(options)`
- `command(request).preview()`
- `command(request).applyAndVerify()`
- `projection().rangeChrome(range)`

Facade rules:

- Facade methods are thin typed wrappers, not a second engine.
- Facade table and selection commands route through `WorkbookCommandService`.
- Facade errors use the same command rejection shape as direct service calls.

## Tables Slice

### Authoritative Model

Use the existing table metadata:

- `name`
- `sheetName`
- `startAddress`
- `endAddress`
- `columnNames`
- `columns`
- `headerRow`
- `totalsRow`
- `style`
- `autoFilter`
- `sortState`

The tables feature plugin owns command descriptors and projection behavior, but it does not introduce a new storage schema.

### Commands

Required table command ids:

- `tables.createFromSelection`
- `tables.upsert`
- `tables.delete`
- `tables.resize`
- `tables.renameHeader`

`tables.createFromSelection` input:

- `range`
- optional `name`
- optional `hasHeaders`
- optional `style`

Behavior:

- Generate a unique table name when `name` is omitted.
- Detect headers from the first selected row by default.
- Canonicalize blank and duplicate headers with Excel-compatible names.
- Produce `upsertTable` preview ops before apply.
- Apply through the existing engine op path.
- Return changed ranges for the full table range, header row, and data body.

`tables.renameHeader` behavior:

- Reuse existing table header rename semantics.
- Preserve structured-reference rewrites for formulas and defined names.
- Return changed ranges for the header cell and dependent table metadata.

`tables.resize` behavior:

- Update `startAddress`, `endAddress`, `columnNames`, and `columns` consistently.
- Preserve table name, style, filter, sort, and totals settings where still valid.
- Reject invalid ranges before mutation.

`tables.delete` behavior:

- Delete table metadata only.
- Do not clear cell values or formulas.
- Return undo ops that restore the table metadata.

### Projection

The tables feature contributes:

- `rangeChrome` for the table rectangle, header row, data body, and totals row.
- `commandMetadata` labels such as `Create table Sales on Sheet1!A1:D20`.
- semantic targets for table name, columns, visible data body, header row, totals row, and changed table range.

Projection output must include provenance:

- `featureId: "tables"`
- `source: "workbook-metadata"`
- `tableName`
- `range`

### Web UI

Add UI-slot support in `apps/web`:

- `toolbar`
- `sidePanel`
- `floatingOverlay`
- `status`

Tables v1 contributes:

- a toolbar action to create a table from the current selection
- a side panel tab showing selected/active table details
- delete and resize actions that route through the command service path

The side panel shows:

- table name
- sheet and range
- columns
- header row enabled
- totals row enabled
- style name or default style
- command status/proof summary for the latest table command

Web rules:

- Keep app-shell changes small.
- Do not put table product logic in `use-worker-workbook-app-state.tsx`.
- Toolbar and side panel use UI slot descriptors and existing Base UI patterns.

### Agent Integration

Existing agent table commands continue to decode and apply through the current bundle shape.

Add semantic preview/readback data:

- table name
- changed table range
- header row range
- data body range
- totals row range when present
- column names and column ranges

Agent preview must continue to use a cloned preview engine and must not mutate the live workbook during preview.

## Implementation Slices

1. Add `@bilig/workbook` feature contracts and tests.
2. Add `@bilig/core` feature registry, projection service, command service, facade, and tests.
3. Add built-in tables feature plugin in core and route table commands through it.
4. Extend agent preview summaries with table semantic targets while keeping existing wire compatibility.
5. Add web UI slots and migrate the tables toolbar/panel UI through the slot path.
6. Add renderer/projection tests proving table chrome is metadata-derived.
7. Run focused tests and `pnpm run ci`.

## Acceptance Criteria

- A caller can register the built-in tables feature and inspect its command descriptors.
- `tables.createFromSelection` previews and applies matching `upsertTable` ops.
- Applying a table command returns undo metadata and can be undone through current engine history.
- Table header rename still rewrites structured references and defined names.
- Table range chrome is available through projection interceptors with metadata provenance.
- Web toolbar can create a table from the current selection.
- Web side panel can inspect the active table and run delete/resize actions.
- Agent previews include table semantic targets without breaking existing command bundles.
- Existing table snapshot, import/export, structured-reference, formula, undo, and sync tests remain green.
- Full CI passes.

## Verification Commands

Run focused checks during implementation:

```sh
pnpm --filter @bilig/workbook test
pnpm --filter @bilig/core test
pnpm --filter @bilig/agent-api test
pnpm --filter @bilig/web test
```

Run the final gate:

```sh
pnpm run ci
```
