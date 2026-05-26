import { mkdir, writeFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'
import type { TemporalWorkPaperQuoteInput, TemporalWorkPaperQuoteResult, TemporalWorkPaperQuoteSummary } from './types'

export async function calculateWorkPaperQuoteActivity(quoteInput: TemporalWorkPaperQuoteInput): Promise<TemporalWorkPaperQuoteResult> {
  const result = calculateWorkPaperQuote(quoteInput)

  await mkdir(dirname(quoteInput.output), { recursive: true })
  await writeFile(quoteInput.output, `${JSON.stringify(result, null, 2)}\n`, 'utf8')

  if (!result.proof.verified) {
    throw new Error(`WorkPaper proof failed: ${JSON.stringify(result.proof)}`)
  }

  return result
}

function calculateWorkPaperQuote(quoteInput: TemporalWorkPaperQuoteInput): TemporalWorkPaperQuoteResult {
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
        temporalBoundary: {
          workflowImportsWorkPaper: false,
          activityOwnsWorkPaper: true,
          payloadShape: 'serializable-patch-and-proof',
        },
        limitations: [
          'This example proves a Temporal Activity boundary with Bilig WorkPaper in normal Node.js code, not an official Temporal sample.',
          'Keep Workflow code deterministic: call this Activity through proxyActivities and keep @bilig/workpaper imports out of Workflow files.',
          'Keep Activity arguments and results compact; store full WorkPaper JSON in an artifact path, object store, or database when payloads grow.',
          'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
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

function readSummary(workbook: WorkPaper, summarySheet: number): TemporalWorkPaperQuoteSummary {
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

function sameSummary(left: TemporalWorkPaperQuoteSummary, right: TemporalWorkPaperQuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
