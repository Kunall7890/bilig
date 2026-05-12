import { WorkPaper } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Revenue: [
      ['Region', 'Customers', 'ARPA', 'MRR'],
      ['West', 20, 1200, '=B2*C2'],
      ['East', 30, 250, '=B3*C3'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Total MRR', '=SUM(Revenue!D2:D3)'],
      ['West Customers', '=Revenue!B2'],
    ],
  },
  {
    maxRows: 1000,
    maxColumns: 64,
    useColumnIndex: true,
  },
)

const summarySheet = requireSheet(workbook, 'Summary')
const summaryRange = {
  start: { sheet: summarySheet, row: 0, col: 0 },
  end: { sheet: summarySheet, row: 2, col: 1 },
}

const output = {
  verified: true,
  range: 'Summary!A1:B3',
  valueReadback: workbook.getRangeValues(summaryRange).map((row) => row.map(readCellValue)),
  serializedReadback: workbook.getRangeSerialized(summaryRange),
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

function readCellValue(cell) {
  if (cell === undefined || cell === null) {
    return null
  }

  if (typeof cell === 'object' && 'value' in cell) {
    return cell.value
  }

  throw new Error(`Expected a scalar WorkPaper cell value, received ${JSON.stringify(cell)}`)
}

function assertOutput(actual) {
  const expected = {
    verified: true,
    range: 'Summary!A1:B3',
    valueReadback: [
      ['Metric', 'Value'],
      ['Total MRR', 31500],
      ['West Customers', 20],
    ],
    serializedReadback: [
      ['Metric', 'Value'],
      ['Total MRR', '=SUM(Revenue!D2:D3)'],
      ['West Customers', '=Revenue!B2'],
    ],
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected range readback result: ${JSON.stringify(actual)}`)
  }
}
