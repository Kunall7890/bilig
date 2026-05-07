# CSV-Shaped Input To WorkPaper

This recipe shows how to turn a small tabular payload into a formula-backed
`@bilig/headless` WorkPaper workbook. It is for service payloads, test fixtures,
and quick evaluator scripts where the source data looks like CSV.

The parser below is intentionally tiny. It handles a simple comma-delimited
fixture with no quoted fields, embedded commas, escaped quotes, multiline cells,
locale-specific number parsing, or import/export metadata. Use a real CSV/XLSX
import pipeline when those features matter.

## Setup

```sh
mkdir bilig-csv-shaped-input
cd bilig-csv-shaped-input
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Create `csv-input.mjs`:

```js
import { WorkPaper } from '@bilig/headless'

const csv = `
region,customers,arpa
West,20,1200
East,30,250
Central,18,300
`.trim()

const rows = parseSimpleCsv(csv)
const revenueRows = rows.map((row, index) => {
  const spreadsheetRow = index + 2
  return [
    row.region,
    readNumber(row.customers, `customers row ${spreadsheetRow}`),
    readNumber(row.arpa, `arpa row ${spreadsheetRow}`),
    `=B${spreadsheetRow}*C${spreadsheetRow}`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Revenue: [['Region', 'Customers', 'ARPA', 'Revenue'], ...revenueRows],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Revenue!D2:D4)'],
    ['West customers', '=SUMIF(Revenue!A2:A4,"West",Revenue!B2:B4)'],
    ['Largest deal', '=MAX(Revenue!D2:D4)'],
  ],
})

const summary = requireSheet(workbook, 'Summary')
const output = {
  totalRevenue: readComputedNumber(workbook, summary, 1, 1, 'Total revenue'),
  westCustomers: readComputedNumber(workbook, summary, 2, 1, 'West customers'),
  largestDeal: readComputedNumber(workbook, summary, 3, 1, 'Largest deal'),
  revenueSheet: workbook.getSheetSerialized(requireSheet(workbook, 'Revenue')),
}

console.log(JSON.stringify(output, null, 2))

function parseSimpleCsv(input) {
  const [headerLine, ...dataLines] = input.split(/\r?\n/)
  const headers = headerLine.split(',').map((header) => header.trim())

  return dataLines.map((line) => {
    const values = line.split(',').map((value) => value.trim())
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? '']))
  })
}

function readNumber(value, label) {
  const parsed = Number(value)
  if (!Number.isFinite(parsed)) {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return parsed
}

function requireSheet(workbook, sheetName) {
  const sheet = workbook.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${sheetName}`)
  }
  return sheet
}

function readComputedNumber(workbook, sheet, row, col, label) {
  const value = workbook.getCellValue({ sheet, row, col })
  if (typeof value !== 'object' || value === null || typeof value.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return Math.round(value.value * 100) / 100
}
```

Run it:

```sh
node csv-input.mjs
```

Expected output:

```json
{
  "totalRevenue": 36900,
  "westCustomers": 20,
  "largestDeal": 24000,
  "revenueSheet": [
    [
      "Region",
      "Customers",
      "ARPA",
      "Revenue"
    ],
    [
      "West",
      20,
      1200,
      "=B2*C2"
    ],
    [
      "East",
      30,
      250,
      "=B3*C3"
    ],
    [
      "Central",
      18,
      300,
      "=B4*C4"
    ]
  ]
}
```

## Notes For Services And Agents

- Treat this as a normalization recipe, not a complete CSV importer.
- Convert untrusted input into the `WorkPaper.buildFromSheets()` array shape
  before creating the workbook.
- Store formulas as formula strings, not precomputed literals, when downstream
  summaries need recalculation.
- Use `getCellValue()` for computed values and `getSheetSerialized()` when you
  need to inspect persisted user inputs.
- Keep import errors explicit. Do not silently coerce missing or non-numeric
  cells into zero unless the business workflow intentionally wants that.

## Validation

For the standalone recipe:

```sh
node csv-input.mjs
```

For a documentation patch in this repository:

```sh
pnpm docs:discovery:check
pnpm run ci
```
