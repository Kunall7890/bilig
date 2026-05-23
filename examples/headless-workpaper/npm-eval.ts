import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from 'bilig-workpaper'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = requireSheet(workbook, 'Inputs')
const summary = requireSheet(workbook, 'Summary')

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const after = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = requireSheet(restored, 'Summary')
const afterRestore = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }))
const nextStep =
  'If this proof matches your workflow, open a concrete blocker or adoption note: https://github.com/proompteng/bilig/discussions/new?category=general'

const output = {
  before,
  after,
  afterRestore,
  sheets: restored.getSheetNames(),
  bytes: saved.length,
  verified: before === 24000 && after === 38400 && afterRestore === 38400,
  nextStep,
}

if (!output.verified) {
  throw new Error(`Unexpected WorkPaper readback: ${JSON.stringify(output)}`)
}

console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper: ReturnType<typeof WorkPaper.buildFromSheets>, sheetName: string): number {
  const sheet = workpaper.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheet
}

function numberValue(cell: unknown): number {
  if (isRecord(cell) && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
