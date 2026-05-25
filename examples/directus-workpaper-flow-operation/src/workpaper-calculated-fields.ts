import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

export interface DirectusCalculatedFieldsInput {
  readonly previousQuantity?: number
  readonly quantity: number
  readonly unitPrice: number
  readonly discountRate?: number
  readonly taxRate?: number
  readonly unitCost?: number
}

export interface DirectusCalculatedFieldsResult {
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
    readonly before: DirectusCalculatedSummary
    readonly after: DirectusCalculatedSummary
    readonly afterRestore: DirectusCalculatedSummary
    readonly persistedDocumentBytes: number
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

export interface DirectusCalculatedSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

export function calculateDirectusWorkPaperFields(rawInput: unknown): DirectusCalculatedFieldsResult {
  const input = parseDirectusCalculatedFieldsInput(rawInput)
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Quantity', input.previousQuantity ?? input.quantity],
      ['Unit price', input.unitPrice],
      ['Discount rate', input.discountRate ?? 0],
      ['Tax rate', input.taxRate ?? 0],
      ['Unit cost', input.unitCost ?? 0],
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

  const inputsSheet = requireSheet(workbook, 'Inputs')
  const summarySheet = requireSheet(workbook, 'Summary')
  const before = readSummary(workbook, summarySheet)

  workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, input.quantity)
  const after = readSummary(workbook, summarySheet)

  const document = exportWorkPaperDocument(workbook, { includeConfig: true })
  const serialized = serializeWorkPaperDocument(document)
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
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
      persistedDocumentBytes: Buffer.byteLength(serialized, 'utf8'),
      verified,
    },
    limitations: [
      'This example proves a Directus Flow operation boundary, not a complete Directus Marketplace package.',
      'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
      'The Directus host must run a Node version supported by @bilig/workpaper.',
    ],
  }
}

function parseDirectusCalculatedFieldsInput(rawInput: unknown): DirectusCalculatedFieldsInput {
  const input = readObject(rawInput, 'operation options')
  return {
    previousQuantity: readOptionalNumber(readProperty(input, 'previousQuantity'), 'previousQuantity'),
    quantity: readNumber(readProperty(input, 'quantity'), 'quantity'),
    unitPrice: readNumber(readProperty(input, 'unitPrice'), 'unitPrice'),
    discountRate: readOptionalNumber(readProperty(input, 'discountRate'), 'discountRate'),
    taxRate: readOptionalNumber(readProperty(input, 'taxRate'), 'taxRate'),
    unitCost: readOptionalNumber(readProperty(input, 'unitCost'), 'unitCost'),
  }
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }
  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): DirectusCalculatedSummary {
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

function sameSummary(left: DirectusCalculatedSummary, right: DirectusCalculatedSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function readObject(value: unknown, label: string): object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${label} must be an object`)
  }
  return value
}

function readProperty(record: object, property: string): unknown {
  return Reflect.get(record, property)
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
