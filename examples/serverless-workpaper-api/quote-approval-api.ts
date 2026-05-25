import { createServer, type IncomingMessage } from 'node:http'
import { pathToFileURL } from 'node:url'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

type QuoteInput = {
  units: number
  listPrice: number
  discount: number
  unitCost: number
  minimumMargin: number
}

type QuoteSummary = {
  listRevenue: number
  discountAmount: number
  netRevenue: number
  totalCost: number
  grossMargin: number
  decision: string
}

type QuoteJsonStorage = {
  loadWorkbookJson(): Promise<string> | string
  saveWorkbookJson(nextWorkbookJson: string): Promise<void> | void
}

type QuoteRequestHandler = (request: Request) => Promise<Response>

const inputCells = {
  units: 'Inputs!B2',
  listPrice: 'Inputs!B3',
  discount: 'Inputs!B4',
  unitCost: 'Inputs!B5',
  minimumMargin: 'Inputs!B6',
} as const

export function createQuoteApprovalStorage(initialWorkbook: WorkPaperInstance = createQuoteApprovalWorkbook()): QuoteJsonStorage {
  let workbookJson = serializeWorkbook(initialWorkbook)
  return {
    async loadWorkbookJson() {
      return workbookJson
    },
    async saveWorkbookJson(nextWorkbookJson: string) {
      workbookJson = nextWorkbookJson
    },
  }
}

export const handleQuoteApprovalRequest = createQuoteApprovalRequestHandler(createQuoteApprovalStorage())

export function createQuoteApprovalRequestHandler(storage: QuoteJsonStorage): QuoteRequestHandler {
  return async function handleStoredQuoteApprovalRequest(request: Request) {
    const url = new URL(request.url)

    if (request.method === 'GET' && url.pathname === '/api/quote/approval') {
      const workbook = await loadWorkbook(storage)
      return json({
        summary: readQuoteSummary(workbook),
        sheets: workbook.getSheetNames(),
      })
    }

    if (request.method === 'POST' && url.pathname === '/api/quote/approval') {
      let quote
      try {
        quote = parseQuoteInput(readJsonRecord(await request.json(), 'request body'))
      } catch (error) {
        return json({ error: error instanceof Error ? error.message : String(error) }, 400)
      }

      const workbook = await loadWorkbook(storage)
      const before = readQuoteSummary(workbook)
      writeQuoteInputs(workbook, quote)
      const after = readQuoteSummary(workbook)
      const workbookJson = serializeWorkbook(workbook)
      await storage.saveWorkbookJson(workbookJson)

      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(workbookJson))
      const restoredSummary = readQuoteSummary(restored)

      return json({
        input: quote,
        inputCells,
        before,
        after,
        restored: restoredSummary,
        checks: {
          decisionChanged: before.decision !== after.decision,
          formulasPersisted: workbookJson.includes('=IF(B6>=Inputs!B6'),
          inputPersisted: JSON.stringify(readQuoteInputs(restored)) === JSON.stringify(quote),
          restoredMatchesAfter: JSON.stringify(restoredSummary) === JSON.stringify(after),
          serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
        },
      })
    }

    return json({ error: 'not found' }, 404)
  }
}

function createQuoteApprovalWorkbook(): WorkPaperInstance {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['List price', 1200],
      ['Discount', 0.1],
      ['Unit cost', 760],
      ['Minimum margin', 0.3],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['List revenue', '=Inputs!B2*Inputs!B3'],
      ['Discount amount', '=B2*Inputs!B4'],
      ['Net revenue', '=B2-B3'],
      ['Total cost', '=Inputs!B2*Inputs!B5'],
      ['Gross margin', '=(B4-B5)/B4'],
      ['Decision', '=IF(B6>=Inputs!B6,"approved","review")'],
    ],
  })
}

function writeQuoteInputs(workbook: WorkPaperInstance, quote: QuoteInput): void {
  const inputs = requireSheet(workbook, 'Inputs')
  workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, quote.units)
  workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, quote.listPrice)
  workbook.setCellContents({ sheet: inputs, row: 3, col: 1 }, quote.discount)
  workbook.setCellContents({ sheet: inputs, row: 4, col: 1 }, quote.unitCost)
  workbook.setCellContents({ sheet: inputs, row: 5, col: 1 }, quote.minimumMargin)
}

async function loadWorkbook(storage: QuoteJsonStorage): Promise<WorkPaperInstance> {
  return createWorkPaperFromDocument(parseWorkPaperDocument(await storage.loadWorkbookJson()))
}

function serializeWorkbook(workbook: WorkPaperInstance): string {
  return serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
}

function parseQuoteInput(record: Record<string, unknown>): QuoteInput {
  return {
    units: readFiniteNumber(record.units, 'units', { min: 1 }),
    listPrice: readFiniteNumber(record.listPrice, 'listPrice', { min: 0 }),
    discount: readFiniteNumber(record.discount, 'discount', { min: 0, max: 0.95 }),
    unitCost: readFiniteNumber(record.unitCost, 'unitCost', { min: 0 }),
    minimumMargin: readFiniteNumber(record.minimumMargin, 'minimumMargin', { min: 0, max: 1 }),
  }
}

function readFiniteNumber(value: unknown, label: string, bounds: { min: number; max?: number }): number {
  const numberValue = Number(value)
  if (!Number.isFinite(numberValue) || numberValue < bounds.min || (bounds.max !== undefined && numberValue > bounds.max)) {
    const maxText = bounds.max === undefined ? '' : ` and <= ${bounds.max.toString()}`
    throw new Error(`${label} must be a finite number >= ${bounds.min.toString()}${maxText}`)
  }
  return numberValue
}

function readQuoteSummary(workbook: WorkPaperInstance): QuoteSummary {
  const summary = requireSheet(workbook, 'Summary')
  return {
    listRevenue: readNumber(workbook, summary, 1, 1, 'List revenue'),
    discountAmount: readNumber(workbook, summary, 2, 1, 'Discount amount'),
    netRevenue: readNumber(workbook, summary, 3, 1, 'Net revenue'),
    totalCost: readNumber(workbook, summary, 4, 1, 'Total cost'),
    grossMargin: readRoundedNumber(workbook, summary, 5, 1, 'Gross margin'),
    decision: readString(workbook, summary, 6, 1, 'Decision'),
  }
}

function readQuoteInputs(workbook: WorkPaperInstance): QuoteInput {
  const inputs = requireSheet(workbook, 'Inputs')
  return {
    units: readCellNumber(workbook, inputs, 1, 1, 'Units'),
    listPrice: readCellNumber(workbook, inputs, 2, 1, 'List price'),
    discount: readCellNumber(workbook, inputs, 3, 1, 'Discount'),
    unitCost: readCellNumber(workbook, inputs, 4, 1, 'Unit cost'),
    minimumMargin: readCellNumber(workbook, inputs, 5, 1, 'Minimum margin'),
  }
}

function requireSheet(workbook: WorkPaperInstance, name: string): number {
  const sheet = workbook.getSheetId(name)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${name}`)
  }
  return sheet
}

function readNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  return Math.round(readCellNumber(workbook, sheet, row, col, label) * 100) / 100
}

function readCellNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (!isJsonRecord(cell) || typeof cell['value'] !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell['value']
}

function readRoundedNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (!isJsonRecord(cell) || typeof cell['value'] !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell['value'] * 10_000) / 10_000
}

function readString(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (!isJsonRecord(cell) || typeof cell['value'] !== 'string') {
    throw new Error(`expected ${label} to be text, received ${JSON.stringify(cell)}`)
  }
  return cell['value']
}

function readJsonRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be a JSON object`)
  }
  return value
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function json(payload: unknown, status = 200): Response {
  return Response.json(payload, {
    status,
    headers: {
      'cache-control': 'no-store',
    },
  })
}

if (process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href) {
  if (process.argv.includes('--serve')) {
    createServer(async (incoming, outgoing) => {
      try {
        const request = await toWebRequest(incoming)
        const response = await handleQuoteApprovalRequest(request)
        outgoing.writeHead(response.status, Object.fromEntries(response.headers))
        outgoing.end(Buffer.from(await response.arrayBuffer()))
      } catch (error) {
        outgoing.writeHead(500, { 'content-type': 'application/json; charset=utf-8' })
        outgoing.end(`${JSON.stringify({ error: error instanceof Error ? error.message : String(error) })}\n`)
      }
    }).listen(8788, () => {
      console.log('Quote approval WorkPaper API listening on http://localhost:8788')
    })
  } else {
    await runSmoke()
  }
}

async function runSmoke(): Promise<void> {
  const storageEvents: string[] = []
  const memoryStorage = createQuoteApprovalStorage()
  const storage: QuoteJsonStorage = {
    async loadWorkbookJson() {
      storageEvents.push('load')
      return memoryStorage.loadWorkbookJson()
    },
    async saveWorkbookJson(nextWorkbookJson: string) {
      storageEvents.push('save')
      await memoryStorage.saveWorkbookJson(nextWorkbookJson)
    },
  }
  const handleRequest = createQuoteApprovalRequestHandler(storage)
  const before = await requestJson(handleRequest, '/api/quote/approval', parseSummaryResponse)
  const edit = await requestJson(handleRequest, '/api/quote/approval', parseEditResponse, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      units: 40,
      listPrice: 1200,
      discount: 0.05,
      unitCost: 760,
      minimumMargin: 0.3,
    }),
  })
  const after = await requestJson(handleRequest, '/api/quote/approval', parseSummaryResponse)
  const output = {
    route: 'Quote approval WorkPaper API',
    inputCells,
    before: before.summary,
    edit,
    after: after.summary,
    verified: true,
    star: 'https://github.com/proompteng/bilig/stargazers',
    watchReleases: 'https://github.com/proompteng/bilig/subscription',
    adoptionBlocker: 'https://github.com/proompteng/bilig/discussions/new?category=general',
    nextStep:
      'If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers',
  }

  assertSmokeOutput(output, storageEvents)
  console.log(JSON.stringify(output, null, 2))
}

async function requestJson<T>(
  handleRequest: QuoteRequestHandler,
  path: string,
  parse: (value: unknown) => T,
  init?: RequestInit,
): Promise<T> {
  const response = await handleRequest(new Request(`http://localhost:8788${path}`, init))
  const body: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`request failed: ${response.status.toString()} ${JSON.stringify(body)}`)
  }
  return parse(body)
}

function parseSummaryResponse(value: unknown): { summary: QuoteSummary } {
  const record = readJsonRecord(value, 'summary response')
  return {
    summary: parseQuoteSummary(record.summary, 'summary response summary'),
  }
}

function parseEditResponse(value: unknown): Record<string, unknown> {
  const record = readJsonRecord(value, 'edit response')
  const checks = readJsonRecord(record.checks, 'edit response checks')
  return {
    input: parseQuoteInput(readJsonRecord(record.input, 'edit input')),
    before: parseQuoteSummary(record.before, 'edit before'),
    after: parseQuoteSummary(record.after, 'edit after'),
    restored: parseQuoteSummary(record.restored, 'edit restored'),
    checks: {
      decisionChanged: readBoolean(checks.decisionChanged, 'decisionChanged'),
      formulasPersisted: readBoolean(checks.formulasPersisted, 'formulasPersisted'),
      inputPersisted: readBoolean(checks.inputPersisted, 'inputPersisted'),
      restoredMatchesAfter: readBoolean(checks.restoredMatchesAfter, 'restoredMatchesAfter'),
      serializedBytes: readPlainNumber(checks.serializedBytes, 'serializedBytes'),
    },
  }
}

function parseQuoteSummary(value: unknown, label: string): QuoteSummary {
  const record = readJsonRecord(value, label)
  return {
    listRevenue: readPlainNumber(record.listRevenue, `${label} listRevenue`),
    discountAmount: readPlainNumber(record.discountAmount, `${label} discountAmount`),
    netRevenue: readPlainNumber(record.netRevenue, `${label} netRevenue`),
    totalCost: readPlainNumber(record.totalCost, `${label} totalCost`),
    grossMargin: readPlainNumber(record.grossMargin, `${label} grossMargin`),
    decision: readPlainString(record.decision, `${label} decision`),
  }
}

function readPlainNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`${label} must be a number`)
  }
  return value
}

function readPlainString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${label} must be a string`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

function assertSmokeOutput(actual: Record<string, unknown>, storageEvents: readonly string[]): void {
  const before = parseQuoteSummary(actual.before, 'output before')
  const after = parseQuoteSummary(actual.after, 'output after')
  const edit = readJsonRecord(actual.edit, 'output edit')
  const checks = readJsonRecord(edit.checks, 'output edit checks')
  const restored = parseQuoteSummary(edit.restored, 'output edit restored')
  const expectedBefore: QuoteSummary = {
    listRevenue: 48000,
    discountAmount: 4800,
    netRevenue: 43200,
    totalCost: 30400,
    grossMargin: 0.2963,
    decision: 'review',
  }
  const expectedAfter: QuoteSummary = {
    listRevenue: 48000,
    discountAmount: 2400,
    netRevenue: 45600,
    totalCost: 30400,
    grossMargin: 0.3333,
    decision: 'approved',
  }

  if (
    JSON.stringify(before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(restored) !== JSON.stringify(expectedAfter) ||
    checks.decisionChanged !== true ||
    checks.formulasPersisted !== true ||
    checks.inputPersisted !== true ||
    checks.restoredMatchesAfter !== true ||
    readPlainNumber(checks.serializedBytes, 'output serializedBytes') <= 0 ||
    !storageEvents.includes('save')
  ) {
    throw new Error(`unexpected quote approval output: ${JSON.stringify(actual)}`)
  }
}

async function toWebRequest(incoming: IncomingMessage): Promise<Request> {
  const origin = `http://${incoming.headers.host ?? 'localhost:8788'}`
  const url = new URL(incoming.url ?? '/', origin)
  const headers = new Headers()

  for (const [name, value] of Object.entries(incoming.headers)) {
    if (Array.isArray(value)) {
      headers.set(name, value.join(', '))
    } else if (value !== undefined) {
      headers.set(name, value)
    }
  }

  const init: RequestInit & { duplex?: 'half' } = {
    method: incoming.method,
    headers,
    body: incoming.method === 'GET' || incoming.method === 'HEAD' ? undefined : await readIncomingBody(incoming),
    duplex: 'half',
  }

  return new Request(url, init)
}

function readIncomingBody(incoming: IncomingMessage): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const chunks: Uint8Array[] = []
    incoming.on('data', (chunk: Buffer | string) => {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk)
    })
    incoming.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')))
    incoming.on('error', reject)
  })
}
