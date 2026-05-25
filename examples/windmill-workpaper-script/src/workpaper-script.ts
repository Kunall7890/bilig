import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

export interface WindmillWorkPaperQuoteResult {
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
    readonly before: WindmillWorkPaperQuoteSummary
    readonly after: WindmillWorkPaperQuoteSummary
    readonly afterRestore: WindmillWorkPaperQuoteSummary
    readonly persistedDocumentBytes: number
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

export interface WindmillWorkPaperQuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

export async function main(
  quantity = 18,
  unitPrice = 125,
  discountRate = 0.1,
  taxRate = 0.08,
  unitCost = 52,
  previousQuantity = 12,
): Promise<WindmillWorkPaperQuoteResult> {
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Quantity', previousQuantity],
      ['Unit price', unitPrice],
      ['Discount rate', discountRate],
      ['Tax rate', taxRate],
      ['Unit cost', unitCost],
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

    workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, quantity)
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
          verified,
        },
        limitations: [
          'This example proves a Windmill TypeScript script boundary, not a Windmill Hub package.',
          'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
          'The Windmill worker must run a JavaScript runtime compatible with @bilig/workpaper.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): WindmillWorkPaperQuoteSummary {
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

function sameSummary(left: WindmillWorkPaperQuoteSummary, right: WindmillWorkPaperQuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
