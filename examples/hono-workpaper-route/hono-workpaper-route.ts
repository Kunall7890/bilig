/**
 * Hono WorkPaper Route Smoke
 *
 * Mounts a WorkPaper request handler inside a Hono app, sends a POST that
 * edits an input cell, and reads back the dependent formula result.
 *
 * The handler uses web-standard Request/Response and is adapted to Hono via
 * c.req.raw — the same pattern shown in docs/node-framework-workpaper-adapters.md.
 *
 * Acceptance proof shape (printed to stdout):
 *
 *   {
 *     "route": "Hono WorkPaper Route",
 *     "inputCell": "Revenue!D2:Dn",
 *     "readbackCell": "Summary!B2",
 *     "before": { "totalRevenue": 36900, "westCustomers": 20, "largestDeal": 24000 },
 *     "edit": {
 *       "records": 4,
 *       "after": { "totalRevenue": 48600, "westCustomers": 20, "largestDeal": 24000 },
 *       "checks": { "totalRevenueChanged": true, "formulasPersisted": true, "serializedBytes": 1194 }
 *     },
 *     "after": { "totalRevenue": 48600, "westCustomers": 20, "largestDeal": 24000 },
 *     "success": true
 *   }
 *
 * Run:
 *
 *   pnpm --dir examples/hono-workpaper-route install --ignore-workspace
 *   pnpm --dir examples/hono-workpaper-route run smoke
 */

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'
import { Hono } from 'hono'

// ---------------------------------------------------------------------------
// Domain types
// ---------------------------------------------------------------------------

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

type RevenueRecord = {
  region: string
  customers: number
  arpa: number
}

type Summary = {
  totalRevenue: number
  westCustomers: number
  largestDeal: number
}

// ---------------------------------------------------------------------------
// Storage
// ---------------------------------------------------------------------------

type WorkPaperJsonStorage = {
  loadWorkbookJson(): Promise<string> | string
  saveWorkbookJson(nextWorkbookJson: string): Promise<void> | void
}

function createInMemoryWorkbookStorage(): WorkPaperJsonStorage {
  let workbookJson = serializeWorkbook(buildRevenueWorkbook(defaultRecords()))
  return {
    loadWorkbookJson() {
      return workbookJson
    },
    saveWorkbookJson(nextWorkbookJson: string) {
      workbookJson = nextWorkbookJson
    },
  }
}

function defaultRecords(): RevenueRecord[] {
  return [
    { region: 'West', customers: 20, arpa: 1200 },
    { region: 'East', customers: 30, arpa: 250 },
    { region: 'Central', customers: 18, arpa: 300 },
  ]
}

// ---------------------------------------------------------------------------
// WorkPaper workbook helpers
// ---------------------------------------------------------------------------

function buildRevenueWorkbook(records: readonly RevenueRecord[]): WorkPaperInstance {
  const dataRows = records.map((record, index) => {
    const row = index + 2
    return [record.region, record.customers, record.arpa, `=B${row}*C${row}`]
  })
  const lastDataRow = records.length + 1

  return WorkPaper.buildFromSheets({
    Revenue: [['Region', 'Customers', 'ARPA', 'Revenue'], ...dataRows],
    Summary: [
      ['Metric', 'Value'],
      ['Total revenue', `=SUM(Revenue!D2:D${lastDataRow})`],
      ['West customers', `=SUMIF(Revenue!A2:A${lastDataRow},"West",Revenue!B2:B${lastDataRow})`],
      ['Largest deal', `=MAX(Revenue!D2:D${lastDataRow})`],
    ],
  })
}

function serializeWorkbook(workbook: WorkPaperInstance): string {
  return serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
}

async function loadWorkbook(storage: WorkPaperJsonStorage): Promise<WorkPaperInstance> {
  return createWorkPaperFromDocument(parseWorkPaperDocument(await storage.loadWorkbookJson()))
}

// ---------------------------------------------------------------------------
// Shared request handler (web-standard Request / Response)
// ---------------------------------------------------------------------------

function createWorkPaperRequestHandler(storage: WorkPaperJsonStorage) {
  return async function handleWorkPaperRequest(request: Request): Promise<Response> {
    const url = new URL(request.url)

    if (request.method === 'GET' && url.pathname === '/api/workpaper/summary') {
      const workbook = await loadWorkbook(storage)
      return jsonResponse({ summary: readSummary(workbook) })
    }

    if (request.method === 'POST' && url.pathname === '/api/workpaper/revenue') {
      let records: RevenueRecord[]
      try {
        const body = await request.json() as { records?: unknown }
        records = normalizeRevenueRecords(body.records)
      } catch (err) {
        return jsonResponse({ error: err instanceof Error ? err.message : String(err) }, 400)
      }

      const before = readSummary(await loadWorkbook(storage))
      const workbook = buildRevenueWorkbook(records)
      const after = readSummary(workbook)
      const workbookJson = serializeWorkbook(workbook)
      await storage.saveWorkbookJson(workbookJson)

      return jsonResponse({
        records: records.length,
        before,
        after,
        checks: {
          totalRevenueChanged: before.totalRevenue !== after.totalRevenue,
          formulasPersisted: workbookJson.includes('=SUM(Revenue!D2:D'),
          serializedBytes: Buffer.byteLength(workbookJson, 'utf8'),
        },
      })
    }

    return jsonResponse({ error: 'not found' }, 404)
  }
}

function jsonResponse(payload: unknown, status = 200): Response {
  return Response.json(payload, { status, headers: { 'cache-control': 'no-store' } })
}

// ---------------------------------------------------------------------------
// Parsers / readers
// ---------------------------------------------------------------------------

function normalizeRevenueRecords(value: unknown): RevenueRecord[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error('records must be a non-empty array')
  }

  return value.map((item, i) => {
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new Error(`record ${i + 1} must be an object`)
    }

    const r = item as Record<string, unknown>
    const region = r.region
    const customers = Number(r.customers)
    const arpa = Number(r.arpa)

    if (typeof region !== 'string' || region.trim() === '') {
      throw new Error(`record ${i + 1} needs a non-empty region`)
    }
    if (!Number.isFinite(customers) || customers < 0) {
      throw new Error(`record ${i + 1} needs non-negative customers`)
    }
    if (!Number.isFinite(arpa) || arpa < 0) {
      throw new Error(`record ${i + 1} needs non-negative arpa`)
    }

    return { region: region.trim(), customers, arpa }
  })
}

function readSummary(workbook: WorkPaperInstance): Summary {
  const summary = requireSheet(workbook, 'Summary')
  return {
    totalRevenue: readNumericCell(workbook, summary, 1, 1, 'Total revenue'),
    westCustomers: readNumericCell(workbook, summary, 2, 1, 'West customers'),
    largestDeal: readNumericCell(workbook, summary, 3, 1, 'Largest deal'),
  }
}

function requireSheet(workbook: WorkPaperInstance, name: string): number {
  const id = workbook.getSheetId(name)
  if (id === undefined) throw new Error(`missing sheet: ${name}`)
  return id
}

function readNumericCell(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workbook.getCellValue({ sheet, row, col })
  if (typeof cell !== 'object' || cell === null || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, got ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

// ---------------------------------------------------------------------------
// Hono adapter (c.req.raw is the web-standard Request)
// ---------------------------------------------------------------------------

const storage = createInMemoryWorkbookStorage()
const handler = createWorkPaperRequestHandler(storage)

/**
 * Hono adapter: pass c.req.raw directly to the shared handler.
 *
 * This is the same two-line pattern documented in
 * docs/node-framework-workpaper-adapters.md:
 *
 *   const honoHandler = createHonoWorkPaperHandler()
 *   app.get('/api/workpaper/summary', honoHandler)
 */
function createHonoWorkPaperHandler() {
  return (c: { req: { raw: Request } }): Promise<Response> => handler(c.req.raw)
}

const honoHandler = createHonoWorkPaperHandler()

const app = new Hono()
app.get('/api/workpaper/summary', honoHandler)
app.post('/api/workpaper/revenue', honoHandler)

// ---------------------------------------------------------------------------
// Smoke: dispatch through app.fetch, parse, assert
// ---------------------------------------------------------------------------

const updatedRecords: RevenueRecord[] = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

const beforeResponse = await fetchJson('GET', '/api/workpaper/summary', parseSummaryResponse)
const editResponse = await fetchJson('POST', '/api/workpaper/revenue', parseEditResponse, { records: updatedRecords })
const afterResponse = await fetchJson('GET', '/api/workpaper/summary', parseSummaryResponse)

const output = {
  route: 'Hono WorkPaper Route',
  inputCell: 'Revenue!D2:Dn',
  readbackCell: 'Summary!B2',
  before: beforeResponse.summary,
  edit: {
    records: editResponse.records,
    after: editResponse.after,
    checks: editResponse.checks,
  },
  after: afterResponse.summary,
  success: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

// ---------------------------------------------------------------------------
// Fetch helper
// ---------------------------------------------------------------------------

async function fetchJson<T>(
  method: 'GET' | 'POST',
  path: string,
  parse: (value: unknown) => T,
  body?: unknown,
): Promise<T> {
  const headers: Record<string, string> = {}
  let requestBody: BodyInit | undefined

  if (body !== undefined) {
    headers['content-type'] = 'application/json'
    requestBody = JSON.stringify(body)
  }

  const response = await app.fetch(new Request(`http://localhost${path}`, { method, headers, body: requestBody }))
  const payload: unknown = await response.json()

  if (!response.ok) {
    throw new Error(`Hono ${method} ${path} → ${response.status}: ${JSON.stringify(payload)}`)
  }

  return parse(payload)
}

// ---------------------------------------------------------------------------
// Response parsers
// ---------------------------------------------------------------------------

type SummaryResponse = { summary: Summary }

type EditResponse = {
  records: number
  before: Summary
  after: Summary
  checks: {
    totalRevenueChanged: boolean
    formulasPersisted: boolean
    serializedBytes: number
  }
}

function parseSummaryResponse(value: unknown): SummaryResponse {
  const r = requireObject(value, 'summary response')
  return { summary: readSummaryFromObject(r.summary, 'summary response.summary') }
}

function parseEditResponse(value: unknown): EditResponse {
  const r = requireObject(value, 'edit response')
  const c = requireObject(r.checks, 'edit response.checks')
  return {
    records: requireNumber(r.records, 'edit response.records'),
    before: readSummaryFromObject(r.before, 'edit response.before'),
    after: readSummaryFromObject(r.after, 'edit response.after'),
    checks: {
      totalRevenueChanged: requireBoolean(c.totalRevenueChanged, 'edit response.checks.totalRevenueChanged'),
      formulasPersisted: requireBoolean(c.formulasPersisted, 'edit response.checks.formulasPersisted'),
      serializedBytes: requireNumber(c.serializedBytes, 'edit response.checks.serializedBytes'),
    },
  }
}

function readSummaryFromObject(value: unknown, label: string): Summary {
  const r = requireObject(value, label)
  return {
    totalRevenue: requireNumber(r.totalRevenue, `${label}.totalRevenue`),
    westCustomers: requireNumber(r.westCustomers, `${label}.westCustomers`),
    largestDeal: requireNumber(r.largestDeal, `${label}.largestDeal`),
  }
}

function requireObject(value: unknown, label: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${label} must be a JSON object, got ${JSON.stringify(value)}`)
  }
  return value as Record<string, unknown>
}

function requireNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') throw new Error(`${label} must be a number, got ${JSON.stringify(value)}`)
  return value
}

function requireBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') throw new Error(`${label} must be a boolean, got ${JSON.stringify(value)}`)
  return value
}

// ---------------------------------------------------------------------------
// Assertions
// ---------------------------------------------------------------------------

function assertOutput(result: typeof output): void {
  const expectedBefore: Summary = { totalRevenue: 36900, westCustomers: 20, largestDeal: 24000 }
  const expectedAfter: Summary = { totalRevenue: 48600, westCustomers: 20, largestDeal: 24000 }

  const failures: string[] = []

  if (JSON.stringify(result.before) !== JSON.stringify(expectedBefore)) {
    failures.push(`before: expected ${JSON.stringify(expectedBefore)}, got ${JSON.stringify(result.before)}`)
  }
  if (JSON.stringify(result.edit.after) !== JSON.stringify(expectedAfter)) {
    failures.push(`edit.after: expected ${JSON.stringify(expectedAfter)}, got ${JSON.stringify(result.edit.after)}`)
  }
  if (JSON.stringify(result.after) !== JSON.stringify(expectedAfter)) {
    failures.push(`after: expected ${JSON.stringify(expectedAfter)}, got ${JSON.stringify(result.after)}`)
  }
  if (result.edit.records !== 4) {
    failures.push(`edit.records: expected 4, got ${result.edit.records}`)
  }
  if (!result.edit.checks.totalRevenueChanged) {
    failures.push('edit.checks.totalRevenueChanged must be true')
  }
  if (!result.edit.checks.formulasPersisted) {
    failures.push('edit.checks.formulasPersisted must be true')
  }
  if (result.edit.checks.serializedBytes <= 0) {
    failures.push('edit.checks.serializedBytes must be positive')
  }

  if (failures.length > 0) {
    throw new Error(`Hono WorkPaper route smoke failed:\n${failures.map((f) => `  \u2022 ${f}`).join('\n')}`)
  }
}
