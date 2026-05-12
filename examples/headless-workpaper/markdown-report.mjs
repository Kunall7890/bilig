import { WorkPaper } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Committed MRR', 39600],
    ['Weighted pipeline MRR', 43400],
    ['Target MRR', 50000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Committed MRR', '=Inputs!B2'],
    ['Weighted pipeline MRR', '=Inputs!B3'],
    ['Target gap', '=Inputs!B4-Inputs!B2'],
  ],
})

const summarySheet = requireSheet(workbook, 'Summary')
const reportRows = [
  ['Committed MRR', readNumber(workbook, summarySheet, 1, 1, 'Committed MRR')],
  ['Weighted pipeline MRR', readNumber(workbook, summarySheet, 2, 1, 'Weighted pipeline MRR')],
  ['Target gap', readNumber(workbook, summarySheet, 3, 1, 'Target gap')],
]
const report = formatMarkdownTable(reportRows)

const output = {
  verified: true,
  report,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function formatMarkdownTable(rows) {
  return ['| Metric | Value |', '| --- | ---: |', ...rows.map(([metric, value]) => `| ${metric} | ${formatCurrency(value)} |`)].join('\n')
}

function formatCurrency(value) {
  return `$${value.toLocaleString('en-US', { maximumFractionDigits: 0 })}`
}

function assertOutput(actual) {
  const expected = {
    verified: true,
    report: [
      '| Metric | Value |',
      '| --- | ---: |',
      '| Committed MRR | $39,600 |',
      '| Weighted pipeline MRR | $43,400 |',
      '| Target gap | $10,400 |',
    ].join('\n'),
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected Markdown report WorkPaper result: ${JSON.stringify(actual)}`)
  }
}
