# @bilig/xlsx-formula-recalc

Diagnose stale cached XLSX formula values in Node and CI, then recalculate the
cells your service actually reads without Excel, LibreOffice, or browser
automation.

This package is for the high-friction Node XLSX workflow: a file library edits
workbook bytes, but the formula cells still carry old cached values. Start with
the cache doctor when you do not know which cells are stale. Use recalculation
after the detector points at the cells that matter.

It fits `xlsx-populate`, SheetJS / `xlsx`, template-generation, GitHub Actions,
and backend file pipelines where stale readback is worse than a hard failure.

## Try The Evaluator First

Run the no-project evaluator demo:

```sh
npx --package @bilig/xlsx-formula-recalc bilig-evaluate --door workbook-compatibility --json
npx --package @bilig/xlsx-formula-recalc bilig-evaluate --door xlsx-cache --json
```

Use `workbook-compatibility` when the question is whether a real workbook has
known integration risks before a Node service or coding agent trusts Bilig with
it. The report lists unsupported functions, external links, VBA payloads,
pivots, volatile functions, stale caches, and concrete `risk.reasons`. It does
not certify Excel compatibility and it does not print a compatibility score.

Expected evaluator shape:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workbook-compatibility",
  "verified": true,
  "evidence": {
    "riskLevel": "high",
    "unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }],
    "volatileFunctions": [{ "name": "NOW", "count": 1 }],
    "formulaCellCount": 3,
    "staleCachedFormulaCount": 2
  }
}
```

Run the report directly against a workbook:

```sh
npx --package @bilig/xlsx-formula-recalc workbook-compatibility-report pricing.xlsx --json
```

## Try The Cache Doctor Directly

Run the no-project demo:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor --demo --json
```

Expected shape:

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "formulaCellCount": 1,
  "inspectedFormulaCellCount": 1,
  "uninspectedFormulaCellCount": 0,
  "staleCachedFormulaCount": 1,
  "cacheStatusSummary": {
    "inspected": 1,
    "stale": 1,
    "fresh": 0,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B2"],
  "formulas": [
    {
      "target": "Summary!B2",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "cacheStatus": "stale",
      "staleCachedValue": true
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true
}
```

The JSON contains proof fields only. It does not include star, release-watch,
or discussion links.

If a coding agent needs to edit workbook state instead of only inspecting XLSX
bytes, do not automate Excel or a browser grid. Run the agent evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Use `bilig-mcp-challenge --json` only when debugging the lower-level MCP
transcript. The [agent MCP evaluator](https://proompteng.github.io/bilig/eval-agent-mcp.html)
shows the same no-key proof path.

Use `cacheStatusSummary` and per-formula `cacheStatus` to separate confirmed
stale caches from missing cached values or formulas without a comparable
recalculated value.

## CI First

Generate a read-only GitHub Actions workflow from npm:

```sh
mkdir -p .github/workflows
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor --print-github-action "**/*.xlsx" \
  > .github/workflows/xlsx-cache-doctor.yml
```

The generated workflow uses `proompteng/bilig@v1`, uploads JSON and Markdown
reports, sets up Node.js 22, pins the npm runtime with `package-version`, and
starts in report-only mode. Add `--fail-on-stale true` when stale formula
caches should block pull requests. For production, pin both the Action ref and
`package-version`; the Action ref pins the wrapper, and `package-version` pins
the npm package it executes.

For a live reviewer path, inspect the
[XLSX Cache Doctor demo PR](https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1).
It runs the Action, finds one stale cached formula value, and uploads the JSON
report artifact.

## If You Arrived From SheetJS or xlsx-populate

`xlsx`, SheetJS-style workbook objects, and `xlsx-populate` are good at file
I/O. They can read workbook bytes, write cells, preserve formulas, and export
an `.xlsx` artifact.

They do not make stale cached formula values fresh inside your Node process.
That is the failure behind issues and searches like:

- `xlsx-populate formula calculated value`
- `SheetJS formula result not updating`
- `xlsx formula recalculation Node.js`
- `get computed value from xlsx formula cell`

Use this package at the file boundary:

1. let your existing library produce XLSX bytes;
2. write them to the XLSX file path your service owns;
3. call `recalculateXlsxFileToFile(...)`;
4. read the proof cells from `result.reads`;
5. return the recalculated output file.

That keeps your current file-writer choice intact and adds only the missing
calculation/readback step.

If you are not sure whether the fix belongs at the XLSX file, SheetJS, ExcelJS,
template, CI, WorkPaper, or agent boundary, start with the
[stale formula readback chooser](https://proompteng.github.io/bilig/stale-formula-readback-chooser.html).

## Install

```sh
npm install @bilig/xlsx-formula-recalc
```

## CLI

If you have a real workbook but do not yet know which formula cells matter,
diagnose it without writing an output file:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor pricing.xlsx --json
```

Inspection imports the workbook, lists formula cells, recomputes every formula
by default, reports stale cached values, and returns suggested `--read` targets
for the recalculation command. If you intentionally pass `--inspect-limit 50`,
the JSON includes the skipped count as `uninspectedFormulaCellCount`.

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "formulaCellCount": 12,
  "inspectedFormulaCellCount": 12,
  "uninspectedFormulaCellCount": 0,
  "inspectionLimit": "all",
  "staleCachedFormulaCount": 3,
  "cacheStatusSummary": {
    "inspected": 12,
    "stale": 3,
    "fresh": 9,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B7"],
  "formulas": [
    {
      "target": "Summary!B7",
      "formula": "=Inputs!B2*Inputs!B3",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "cacheStatus": "stale",
      "staleCachedValue": true
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

`xlsx-cache-doctor` is a readable alias for
`xlsx-recalc pricing.xlsx --inspect --json`. Use it for issue triage, CI, and
pull-request checks when the only question is whether committed XLSX files have
stale cached formula values.

When you know which cells matter, run the recalculation check:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json
```

That command creates a tiny workbook, changes `Inputs!B2` and `Inputs!B3`,
recalculates `Summary!B2`, writes `bilig-formula-recalc-demo.xlsx`, and prints
the recalculated value.

For an existing workbook:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes a recalculated workbook and prints readback values. Cell targets
must be sheet-qualified A1 references such as `Inputs!B2` or
`'Pricing Model'!F12`.

For workbooks with external links, pass companion workbook files so cached link
values can be refreshed before recalculation:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc model.xlsx \
  --external-workbook rates.xlsx \
  --read Model!C1 \
  --out model.recalculated.xlsx \
  --json
```

When the link target in the workbook is an exact path or URI that does not match
the local companion filename, bind the companion explicitly:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc model.xlsx \
  --external-workbook-target ./fixtures/rates-current.xlsx file:///tmp/rates.xlsx \
  --read Model!C1 \
  --json
```

Ambiguous companion matches fail closed: the command preserves existing
external-link cache values, emits a warning, and includes hydration diagnostics
in JSON output.

For a maintained external-workbook proof with companion hydration diagnostics,
run
[external workbook recalculation proof in Node.js](https://proompteng.github.io/bilig/external-workbook-recalc-proof.html).

## API

Use `inspectXlsxCacheFile` when a service or test runner needs the
cache-doctor report without shelling out to the CLI. It stays on the same
file-backed streaming-native path as `xlsx-cache-doctor`:

```ts
import { inspectXlsxCacheFile } from '@bilig/xlsx-formula-recalc'

const report = await inspectXlsxCacheFile('pricing.xlsx', {
  maxRssBytes: 350 * 1024 * 1024,
})

if (report.staleCachedFormulaCount > 0) {
  throw new Error(
    report.formulas
      .filter((formula) => formula.cacheStatus === 'stale')
      .map((formula) => formula.target)
      .join(', '),
  )
}
```

The API returns the same `schemaVersion`, `cacheStatusSummary`, per-formula
`cacheStatus`, and `suggestedReads` fields as the JSON CLI report.

```ts
import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('pricing.xlsx', {
  outputPath: 'pricing.recalculated.xlsx',
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
  engine: 'streaming-native',
})

console.log(result.reads['Summary!B7'])
```

External companion workbooks stay on the native file-to-file path:

```ts
import { readFile } from 'node:fs/promises'

import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('model.xlsx', {
  outputPath: 'model.recalculated.xlsx',
  externalWorkbooks: [
    {
      bytes: await readFile('rates.xlsx'),
      fileName: 'rates.xlsx',
      target: 'file:///tmp/rates.xlsx',
    },
  ],
  reads: ['Model!C1'],
  engine: 'streaming-native',
})

console.log(result.diagnostics?.externalWorkbookHydration)
```

If another library already produced workbook bytes instead of a file path, use
the explicit legacy compatibility import:

```ts
import { recalculateXlsx } from '@bilig/workpaper/xlsx'

const output = await workbook.outputAsync('nodebuffer') // for example, from xlsx-populate

const result = recalculateXlsx(output, {
  reads: ['Summary!B7'],
})
```

For the full workbook API, use `@bilig/workpaper`. For the old bytes-in,
bytes-out compatibility API, import from
`@bilig/workpaper/xlsx` and install `@bilig/workpaper`
explicitly.

## Common Boundaries

| Existing tool                       | Keep using it for                                      | Add this package when                            |
| ----------------------------------- | ------------------------------------------------------ | ------------------------------------------------ |
| `xlsx-populate`                     | template editing and workbook generation               | formula cells need fresh cached values in Node   |
| SheetJS / `xlsx`                    | broad XLSX parsing, writing, and file interchange      | edited inputs must update dependent formulas now |
| ExcelJS                             | styled reports, sheets, tables, and ExcelJS workbooks  | use `@bilig/exceljs-formula-recalc`              |
| Excel, LibreOffice, Microsoft Graph | exact spreadsheet application behavior                 | you cannot depend on an external app or API call |
| `@bilig/workpaper`                  | service-owned formula workbook state with JSON storage | the workbook does not have to stay XLSX-first    |

## After The Proof

If the recalculation proof matches the Node/XLSX workflow you need, keep the
repository and release feed nearby:
<https://github.com/proompteng/bilig>.

Watch releases if the package is close to a production path:
<https://github.com/proompteng/bilig/subscription>.

If a real workbook edge case blocks adoption, open the smallest reduced
workbook blocker:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Scope

Use this when a Node service needs deterministic formula readback after it
changes XLSX inputs. It is not a full Excel clone: unsupported Excel functions,
external workbook links, macros, and volatile functions may need review. Import
warnings are returned in `result.warnings`.

Full docs: <https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html>

Compatibility note: the unscoped `xlsx-formula-recalc` package remains
published as a search and migration alias. Use `@bilig/xlsx-formula-recalc` in
new projects.
