import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Deals: [
    ['Region', 'Segment', 'Customers', 'ARPA', 'Revenue'],
    ['West', 'Enterprise', 12, 1200, '=C2*D2'],
    ['East', 'SMB', 30, 250, '=C3*D3'],
    ['West', 'SMB', 18, 300, '=C4*D4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Deals!E2:E4)'],
    ['West customers', '=SUMIF(Deals!A2:A4,"West",Deals!C2:C4)'],
    ['Enterprise ARPA', '=XLOOKUP("Enterprise",Deals!B2:B4,Deals!D2:D4)'],
    ['Qualified customer counts', '=FILTER(Deals!C2:C4,Deals!C2:C4>=18)'],
  ],
})

const dealsSheet = requireSheet(workbook, 'Deals')
const summarySheet = requireSheet(workbook, 'Summary')

const initial = {
  totalRevenue: readNumber(workbook, summarySheet, 1, 1, 'initial total revenue'),
  westCustomers: readNumber(workbook, summarySheet, 2, 1, 'initial west customers'),
}

workbook.batch(() => {
  workbook.setCellContents({ sheet: dealsSheet, row: 1, col: 2 }, 20)
})

const persisted = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(persisted))
const restoredSummarySheet = requireSheet(restored, 'Summary')

const afterAgentEdit = {
  totalRevenue: readNumber(restored, restoredSummarySheet, 1, 1, 'edited total revenue'),
  westCustomers: readNumber(restored, restoredSummarySheet, 2, 1, 'edited west customers'),
  enterpriseArpa: readNumber(restored, restoredSummarySheet, 3, 1, 'enterprise ARPA'),
  qualifiedCustomerCounts: readNumberColumn(restored, restoredSummarySheet, 4, 6, 1, 'qualified customer count'),
}

const output = {
  initial,
  afterAgentEdit,
  persistedSheets: restored.getSheetNames(),
}

assertSummary(output)
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
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readNumberColumn(workpaper, sheet, startRow, endRow, col, label) {
  return workpaper
    .getRangeValues({
      start: { sheet, row: startRow, col },
      end: { sheet, row: endRow, col },
    })
    .flat()
    .map((cell, index) => {
      if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
        throw new Error(`Expected ${label} ${index + 1} to be a number, received ${JSON.stringify(cell)}`)
      }
      return cell.value
    })
}

function assertSummary(summary) {
  const expected = {
    initial: {
      totalRevenue: 27300,
      westCustomers: 30,
    },
    afterAgentEdit: {
      totalRevenue: 36900,
      westCustomers: 38,
      enterpriseArpa: 1200,
      qualifiedCustomerCounts: [20, 30, 18],
    },
    persistedSheets: ['Deals', 'Summary'],
  }

  if (JSON.stringify(summary) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected WorkPaper result: ${JSON.stringify(summary)}`)
  }
}
