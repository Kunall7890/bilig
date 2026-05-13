import { WorkPaper } from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const csvInput = [
  ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
  ['Widget A', 100, 150, 200, 250],
  ['Widget B', 80, 90, 100, 110],
  ['Widget C', 300, 310, 320, 330],
]

const workbook = WorkPaper.buildFromSheets({
  Data: csvInput,
})

const dataSheet = requireSheet(workbook, 'Data')

// Add one formula-backed summary cell
workbook.setCellContents({ sheet: dataSheet, row: 4, col: 0 }, 'Total Q1')
workbook.setCellContents({ sheet: dataSheet, row: 4, col: 1 }, '=SUM(B2:B4)')

// Read the display/result back
const output = {
  success: true,
  totalQ1: readNumber(workbook, dataSheet, 4, 1, 'total Q1'),
}

assertSummary(output)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertSummary(summary: typeof output): void {
  const expected = {
    success: true,
    totalQ1: 480,
  }

  if (JSON.stringify(summary) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected WorkPaper result: ${JSON.stringify(summary)}`)
  }
}
