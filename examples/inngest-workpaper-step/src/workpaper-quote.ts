import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

export interface InngestWorkPaperQuoteInput {
  readonly previousQuantity?: number
  readonly quantity: number
  readonly unitPrice: number
  readonly discountRate?: number
  readonly taxRate?: number
  readonly unitCost?: number
}

export interface InngestWorkPaperQuoteResult {
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
    readonly before: InngestWorkPaperQuoteSummary
    readonly after: InngestWorkPaperQuoteSummary
    readonly afterRestore: InngestWorkPaperQuoteSummary
    readonly persistedDocumentBytes: number
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

export interface InngestWorkPaperQuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

export function calculateWorkPaperQuote(rawInput: InngestWorkPaperQuoteInput): InngestWorkPaperQuoteResult {
  const input = normalizeInput(rawInput)
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Quantity', input.previousQuantity],
      ['Unit price', input.unitPrice],
      ['Discount rate', input.discountRate],
      ['Tax rate', input.taxRate],
      ['Unit cost', input.unitCost],
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

    workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, input.quantity)
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
          'This example proves an Inngest step.run boundary, not an official Inngest template.',
          'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
          'The Inngest worker must run a JavaScript runtime compatible with @bilig/workpaper.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function normalizeInput(input: InngestWorkPaperQuoteInput): Required<InngestWorkPaperQuoteInput> {
  const quantity = readNumber(input.quantity, 'quantity')

  return {
    previousQuantity: readOptionalNumber(input.previousQuantity, 'previousQuantity') ?? quantity,
    quantity,
    unitPrice: readNumber(input.unitPrice, 'unitPrice'),
    discountRate: readOptionalNumber(input.discountRate, 'discountRate') ?? 0,
    taxRate: readOptionalNumber(input.taxRate, 'taxRate') ?? 0,
    unitCost: readOptionalNumber(input.unitCost, 'unitCost') ?? 0,
  }
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): InngestWorkPaperQuoteSummary {
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

function sameSummary(left: InngestWorkPaperQuoteSummary, right: InngestWorkPaperQuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${label} must be a finite number`)
  }

  return value
}

function readOptionalNumber(value: unknown, label: string): number | undefined {
  if (value === undefined || value === null || value === '') {
    return undefined
  }

  return readNumber(value, label)
}
