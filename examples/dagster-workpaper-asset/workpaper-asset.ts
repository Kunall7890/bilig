import { appendFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import { inflateSync } from 'node:zlib'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

interface DagsterWorkPaperQuoteInput {
  readonly previousQuantity: number
  readonly quantity: number
  readonly unitPrice: number
  readonly discountRate: number
  readonly taxRate: number
  readonly unitCost: number
  readonly output: string
}

interface DagsterWorkPaperQuoteSummary {
  readonly subtotal: number
  readonly discountAmount: number
  readonly taxableAmount: number
  readonly taxAmount: number
  readonly total: number
  readonly marginAmount: number
}

interface DagsterWorkPaperQuoteResult {
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
    readonly before: DagsterWorkPaperQuoteSummary
    readonly after: DagsterWorkPaperQuoteSummary
    readonly afterRestore: DagsterWorkPaperQuoteSummary
    readonly persistedDocumentBytes: number
    readonly outputFile: string
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

interface DagsterPipesContext {
  readonly asset_keys?: readonly unknown[]
  readonly extras?: Record<string, unknown>
}

const dagsterContext = await readDagsterPipesContext()
const cliInput = readCliInput(process.argv.slice(2), dagsterContext?.extras)
const result = calculateWorkPaperQuote(cliInput)

await mkdir(dirname(cliInput.output), { recursive: true })
await writeFile(cliInput.output, `${JSON.stringify(result, null, 2)}\n`, 'utf8')
await reportDagsterMaterialization(dagsterContext, result)

console.log(JSON.stringify(result, null, 2))

if (!result.proof.verified) {
  throw new Error(`WorkPaper proof failed: ${JSON.stringify(result.proof)}`)
}

function calculateWorkPaperQuote(quoteInput: DagsterWorkPaperQuoteInput): DagsterWorkPaperQuoteResult {
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
          'This example proves a Dagster asset boundary with a Node WorkPaper subprocess, not an official Dagster example.',
          'Keep Excel or another oracle in the loop for macros, pivots, external links, and exact desktop Excel behavior.',
          'Keep WorkPaper inputs and proof artifacts small enough for Dagster metadata; store full documents in an artifact path or object store.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

async function readDagsterPipesContext(): Promise<DagsterPipesContext | undefined> {
  const decodedParam = decodeDagsterPipesParam(process.env['DAGSTER_PIPES_CONTEXT'])

  if (decodedParam === undefined) {
    return undefined
  }

  const path = readPathFromDagsterPipesParam(decodedParam)
  return parseDagsterPipesContext(JSON.parse(await readFile(path, 'utf8')))
}

async function reportDagsterMaterialization(
  context: DagsterPipesContext | undefined,
  quoteResult: DagsterWorkPaperQuoteResult,
): Promise<void> {
  const decodedParam = decodeDagsterPipesParam(process.env['DAGSTER_PIPES_MESSAGES'])

  if (context === undefined || decodedParam === undefined) {
    return
  }

  const path = readPathFromDagsterPipesParam(decodedParam)
  await appendFile(
    path,
    `${JSON.stringify({
      method: 'report_asset_materialization',
      params: {
        asset_key: firstAssetKey(context),
        data_version: null,
        metadata: {
          workpaper_patch: { raw_value: quoteResult.patch, type: 'json' },
          workpaper_proof: { raw_value: quoteResult.proof, type: 'json' },
          proof_file: { raw_value: quoteResult.proof.outputFile, type: 'path' },
          edited_cell: { raw_value: quoteResult.proof.editedCell, type: 'text' },
          total: { raw_value: quoteResult.patch.total, type: 'float' },
        },
      },
    })}\n`,
    'utf8',
  )
}

function decodeDagsterPipesParam(value: string | undefined): unknown {
  if (value === undefined || value.length === 0) {
    return undefined
  }

  return JSON.parse(inflateSync(Buffer.from(value, 'base64')).toString('utf8'))
}

function readPathFromDagsterPipesParam(value: unknown): string {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`Dagster Pipes parameter must be an object, got ${JSON.stringify(value)}`)
  }

  const path = Reflect.get(value, 'path')
  if (typeof path !== 'string' || path.length === 0) {
    throw new Error(`Dagster Pipes parameter must include a path, got ${JSON.stringify(value)}`)
  }

  return path
}

function parseDagsterPipesContext(value: unknown): DagsterPipesContext {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`Dagster Pipes context must be an object, got ${JSON.stringify(value)}`)
  }

  const assetKeys = Reflect.get(value, 'asset_keys')
  const extras = Reflect.get(value, 'extras')

  if (assetKeys !== undefined && !Array.isArray(assetKeys)) {
    throw new Error(`Dagster Pipes context asset_keys must be an array, got ${JSON.stringify(assetKeys)}`)
  }

  if ((extras !== undefined && typeof extras !== 'object') || extras === null || Array.isArray(extras)) {
    throw new Error(`Dagster Pipes context extras must be an object, got ${JSON.stringify(extras)}`)
  }

  return {
    asset_keys: assetKeys,
    extras,
  }
}

function firstAssetKey(context: DagsterPipesContext): unknown {
  return Array.isArray(context.asset_keys) && context.asset_keys.length > 0 ? context.asset_keys[0] : ['bilig_workpaper_quote_asset']
}

function readCliInput(args: readonly string[], extras: Record<string, unknown> | undefined): DagsterWorkPaperQuoteInput {
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
    previousQuantity: readNumberInput(values.get('previous-quantity'), readExtra(extras, 'previous_quantity'), '12', 'previous quantity'),
    quantity: readNumberInput(values.get('quantity'), readExtra(extras, 'quantity'), '18', 'quantity'),
    unitPrice: readNumberInput(values.get('unit-price'), readExtra(extras, 'unit_price'), '125', 'unit price'),
    discountRate: readNumberInput(values.get('discount-rate'), readExtra(extras, 'discount_rate'), '0.1', 'discount rate'),
    taxRate: readNumberInput(values.get('tax-rate'), readExtra(extras, 'tax_rate'), '0.08', 'tax rate'),
    unitCost: readNumberInput(values.get('unit-cost'), readExtra(extras, 'unit_cost'), '52', 'unit cost'),
    output: values.get('output') ?? 'workpaper-proof.json',
  }
}

function readExtra(extras: Record<string, unknown> | undefined, key: string): unknown {
  return extras === undefined ? undefined : Reflect.get(extras, key)
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): DagsterWorkPaperQuoteSummary {
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

function readNumberInput(cliValue: string | undefined, extraValue: unknown, fallback: string, label: string): number {
  if (cliValue !== undefined) {
    return readNumber(cliValue, label)
  }

  if (extraValue !== undefined) {
    if (typeof extraValue !== 'number' && typeof extraValue !== 'string') {
      throw new Error(`${label} must be a finite number`)
    }

    return readNumber(String(extraValue), label)
  }

  return readNumber(fallback, label)
}

function readNumber(value: string, label: string): number {
  const numeric = Number(value)

  if (!Number.isFinite(numeric)) {
    throw new Error(`${label} must be a finite number`)
  }

  return numeric
}

function sameSummary(left: DagsterWorkPaperQuoteSummary, right: DagsterWorkPaperQuoteSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
