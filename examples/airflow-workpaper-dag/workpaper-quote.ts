import { mkdir, writeFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

interface AirflowWorkPaperQuoteInput {
  readonly previousQuantity: number
  readonly quantity: number
  readonly unitPrice: number
  readonly discountRate: number
  readonly taxRate: number
  readonly unitCost: number
  readonly output: string
}

interface AirflowWorkPaperQuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

interface AirflowWorkPaperQuoteResult {
  readonly patch: {
    readonly subtotal: number
    readonly discount_amount: number
    readonly taxable_amount: number
    readonly tax_amount: number
    readonly total: number
    readonly margin_amount: number
  }
  readonly proof: {
    readonly editedCell: 'Inputs!B2'
    readonly before: AirflowWorkPaperQuoteSummary
    readonly after: AirflowWorkPaperQuoteSummary
    readonly afterRestore: AirflowWorkPaperQuoteSummary
    readonly persistedDocumentBytes: number
    readonly outputFile: string
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

const cliInput = readCliInput(process.argv.slice(2))
const result = calculateWorkPaperQuote(cliInput)

await mkdir(dirname(cliInput.output), { recursive: true })
await writeFile(cliInput.output, `${JSON.stringify(result, null, 2)}\n`, 'utf8')

console.log(JSON.stringify(result, null, 2))

if (!result.proof.verified) {
  throw new Error(`WorkPaper proof failed: ${JSON.stringify(result.proof)}`)
}

function calculateWorkPaperQuote(quoteInput: AirflowWorkPaperQuoteInput): AirflowWorkPaperQuoteResult {
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Quantity', quoteInput.previousQuantity],
      ['Unit price', quoteInput.unitPrice],
      ['Discount rate', quoteInput.discountRate],
      ['Tax rate', quoteInput.taxRate],
      ['Unit cost', quoteInput.unitCost],
    ],
    Summary: [
      ['Field', 'Value'],
      ['Subtotal', '=Inputs!B2*Inputs!B3'],
      ['Discount amount', '=B2*Inputs!B4'],
      ['Taxable amount', '=B2-B3'],
      ['Tax amount', '=B4*Inputs!B5'],
      ['Total', '=B4+B5'],
      ['Margin amount', '=B4-(Inputs!B2*Inputs!B6)'],
    ],
  })

  try {
    const inputsSheet = requireSheet(workbook, 'Inputs')
    const summarySheet = requireSheet(workbook, 'Summary')
    const before = readSummary(workbook, summarySheet)

    workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, quoteInput.quantity)
    const after = readSummary(workbook, summarySheet)

    const document = exportWorkPaperDocument(workbook, { includeConfig: true })
    const serialized = serializeWorkPaperDocument(document)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))

    try {
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const afterRestore = readSummary(restored, restoredSummarySheet)
      const verified = sameSummary(after, afterRestore) && serialized.length > 0

      return {
        patch: {
          subtotal: after.subtotal,
          discount_amount: after.discountAmount,
          taxable_amount: after.taxableAmount,
          tax_amount: after.taxAmount,
          total: after.total,
          margin_amount: after.marginAmount,
        },
        proof: {
          editedCell: 'Inputs!B2',
          before,
          after,
          afterRestore,
          persistedDocumentBytes: new TextEncoder().encode(serialized).byteLength,
          outputFile: quoteInput.output,
          verified,
        },
        limitations: [
          'This example proves an Apache Airflow DAG boundary with a Node WorkPaper step, not an official Apache Airflow example.',
          'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
          'The Airflow worker image or environment must provide Node plus npm dependencies compatible with @bilig/workpaper.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function readCliInput(args: readonly string[]): AirflowWorkPaperQuoteInput {
  const values = new Map<string, string>()

  for (let index = 0; index < args.length; index += 1) {
    const key = args[index]

    if (!key?.startsWith('--')) {
      continue
    }

    const next = args[index + 1]
    if (next === undefined) {
      throw new Error(`Missing value for ${key}`)
    }

    values.set(key.slice(2), next)
    index += 1
  }

  return {
    previousQuantity: readNumber(values.get('previous-quantity') ?? '12', 'previous quantity'),
    quantity: readNumber(values.get('quantity') ?? '18', 'quantity'),
    unitPrice: readNumber(values.get('unit-price') ?? '125', 'unit price'),
    discountRate: readNumber(values.get('discount-rate') ?? '0.1', 'discount rate'),
    taxRate: readNumber(values.get('tax-rate') ?? '0.08', 'tax rate'),
    unitCost: readNumber(values.get('unit-cost') ?? '52', 'unit cost'),
    output: values.get('output') ?? 'workpaper-proof.json',
  }
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): AirflowWorkPaperQuoteSummary {
  return {
    subtotal: readNumberCell(workbook, summarySheet, 1, 'subtotal'),
    discountAmount: readNumberCell(workbook, summarySheet, 2, 'discount amount'),
    taxableAmount: readNumberCell(workbook, summarySheet, 3, 'taxable amount'),
    taxAmount: readNumberCell(workbook, summarySheet, 4, 'tax amount'),
    total: readNumberCell(workbook, summarySheet, 5, 'total'),
    marginAmount: readNumberCell(workbook, summarySheet, 6, 'margin amount'),
  }
}

function readNumberCell(workbook: WorkPaper, sheet: number, row: number, label: string): number {
  const cell = workbook.getCellValue({ sheet, row, col: 1 })
  const value = typeof cell === 'object' && cell !== null ? Reflect.get(cell, 'value') : cell

  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected numeric ${label}, got ${JSON.stringify(cell)}`)
  }

  return value
}

function readNumber(value: string, label: string): number {
  const numeric = Number(value)

  if (!Number.isFinite(numeric)) {
    throw new Error(`${label} must be a finite number`)
  }

  return numeric
}

function sameSummary(left: AirflowWorkPaperQuoteSummary, right: AirflowWorkPaperQuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
