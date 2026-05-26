# xlsx-formula-recalc

Recalculate XLSX formula values in Node.js after SheetJS, ExcelJS, or
xlsx-populate edits without opening Excel, LibreOffice, or a browser.

This package is a narrow wrapper around Bilig WorkPaper for the high-friction Node XLSX workflow:

1. import an XLSX workbook,
2. edit input cells,
3. recalculate formulas,
4. read proof values,
5. export an updated XLSX.

It fits the common `xlsx-populate`, SheetJS, and template-generation case
where the file writer can create or edit the XLSX, but the Node service also
needs fresh formula readback before returning.

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
2. call `recalculateXlsx(...)`;
3. read the proof cells from `result.reads`;
4. write `result.xlsx` if the recalculated workbook artifact is needed.

That keeps your current file-writer choice intact and adds only the missing
calculation/readback step.

For a cross-library proof, run
[`examples/recalc-bridge-workflows`](../../examples/recalc-bridge-workflows).
It edits the same workbook through SheetJS/`xlsx`, `xlsx-populate`, and
ExcelJS, then verifies that Bilig refreshes the stale formula result.

For the SheetJS-specific boundary, read
[SheetJS formula result not updating in Node.js](../../docs/sheetjs-formula-result-not-updating-node.md).
The SheetJS-named `sheetjs-formula-recalc` package is also published for teams
that search and install through the SheetJS / `xlsx` pipeline. It uses the same
underlying recalculation implementation as this package.

```sh
npx --package sheetjs-formula-recalc sheetjs-recalc --demo --json
```

## Install

```sh
npm install xlsx-formula-recalc
```

## CLI

Run a self-contained proof first:

```sh
npx --package xlsx-formula-recalc xlsx-recalc --demo --json
```

That command creates a tiny workbook, changes `Inputs!B2` and `Inputs!B3`,
recalculates `Summary!B2`, writes `bilig-formula-recalc-demo.xlsx`, and prints
a proof object with `verified: true`, the recalculated value, and explicit next
actions:

```json
{
  "reads": {
    "Summary!B2": {
      "value": 72000
    }
  },
  "warnings": [],
  "verified": true,
  "star": "https://github.com/proompteng/bilig/stargazers",
  "watchReleases": "https://github.com/proompteng/bilig/subscription",
  "adoptionBlocker": "https://github.com/proompteng/bilig/discussions/new?category=general",
  "nextStep": "If this XLSX recalculation proof matches your workflow, star or bookmark Bilig; if it almost works, open the concrete workbook blocker."
}
```

Keep the proof first: use the star or release-watch links only after the
recalculated value and warnings match the workflow you are evaluating.

For an existing workbook:

```sh
npx --package xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes a recalculated workbook and prints readback values. Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or `'Pricing Model'!F12`.

For workbooks with external links, pass companion workbook files so cached link
values can be refreshed before recalculation:

```sh
npx --package xlsx-formula-recalc xlsx-recalc model.xlsx \
  --external-workbook rates.xlsx \
  --read Model!C1 \
  --out model.recalculated.xlsx \
  --json
```

When the link target in the workbook is an exact path or URI that does not match
the local companion filename, bind the companion explicitly:

```sh
npx --package xlsx-formula-recalc xlsx-recalc model.xlsx \
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

```ts
import { recalculateXlsx } from 'xlsx-formula-recalc'

const result = recalculateXlsx(await fs.promises.readFile('pricing.xlsx'), {
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

await fs.promises.writeFile('pricing.recalculated.xlsx', result.xlsx)
console.log(result.reads['Summary!B7'])
```

External companion workbooks use the same matching rules as the CLI:

```ts
const result = recalculateXlsx(await fs.promises.readFile('model.xlsx'), {
  externalWorkbooks: [
    {
      bytes: await fs.promises.readFile('rates.xlsx'),
      fileName: 'rates.xlsx',
      target: 'file:///tmp/rates.xlsx',
    },
  ],
  reads: ['Model!C1'],
})

console.log(result.diagnostics?.externalWorkbookHydration)
```

If another library already produced the workbook bytes, pass those bytes directly:

```ts
const output = await workbook.outputAsync('nodebuffer') // for example, from xlsx-populate

const result = recalculateXlsx(output, {
  reads: ['Summary!B7'],
})
```

For the full workbook API, import `WorkPaper`, `importXlsx`, and `exportXlsx` from this package.

## Common Boundaries

| Existing tool                          | Keep using it for                                      | Add this package when                               |
| -------------------------------------- | ------------------------------------------------------ | --------------------------------------------------- |
| `xlsx-populate`                        | template editing and workbook generation               | formula cells need fresh cached values in Node      |
| SheetJS / `xlsx`                       | broad XLSX parsing, writing, and file interchange      | edited inputs must update dependent formulas now    |
| ExcelJS                                | styled reports, sheets, tables, and ExcelJS workbooks  | use `exceljs-formula-recalc` for the ExcelJS object |
| Excel, LibreOffice, Microsoft Graph    | exact spreadsheet application behavior                 | you cannot depend on an external app or API call    |
| `@bilig/headless` or `bilig-workpaper` | service-owned formula workbook state with JSON storage | the workbook does not have to stay XLSX-first       |

## Scope

Use this when a Node service needs deterministic formula readback after it changes XLSX inputs. It is not a full Excel clone: unsupported Excel functions, external workbook links, macros, and volatile functions may need review. Import warnings are returned in `result.warnings`.
