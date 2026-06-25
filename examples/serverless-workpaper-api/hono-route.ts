import { Hono } from 'hono'

import { createHonoWorkPaperHandler } from './framework-adapters.ts'
import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler } from './route.ts'

type Summary = {
  largestDeal: number
  totalRevenue: number
  westCustomers: number
}

type EditChecks = {
  formulasPersisted: boolean
  serializedBytes: number
  totalRevenueChanged: boolean
}

type EditResult = {
  after: Summary
  checks: EditChecks
  records: number
}

type HonoRouteProof = {
  after: Summary
  before: Summary
  edit: EditResult
  inputCell: string
  readbackCell: string
  route: string
  success: boolean
}

const app = new Hono()
const handler = createWorkPaperRequestHandler(createInMemoryWorkbookStorage())
const honoHandler = createHonoWorkPaperHandler(handler)

app.get('/api/workpaper/summary', honoHandler)
app.post('/api/workpaper/revenue', honoHandler)

const updatedRevenueRecords = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

const before = readSummaryPayload(await requestJson('GET', '/api/workpaper/summary'), 'before')
const edit = readEditPayload(
  await requestJson('POST', '/api/workpaper/revenue', {
    records: updatedRevenueRecords,
  }),
)
const after = readSummaryPayload(await requestJson('GET', '/api/workpaper/summary'), 'after')

const proof: HonoRouteProof = {
  route: 'Hono WorkPaper Route',
  inputCell: 'Revenue!A2:D5',
  readbackCell: 'Summary!B2',
  before,
  edit,
  after,
  success: true,
}

assertProof(proof)
console.log(JSON.stringify(proof, null, 2))

async function requestJson(method: 'GET' | 'POST', path: string, body?: unknown): Promise<unknown> {
  const headers = new Headers()
  const init: RequestInit = {
    method,
    headers,
  }

  if (body !== undefined) {
    if (method === 'GET') {
      throw new Error('GET requests must not include a body')
    }
    headers.set('content-type', 'application/json')
    init.body = JSON.stringify(body)
  }

  const response = await app.fetch(new Request(`http://localhost${path}`, init))
  const payload: unknown = await response.json()

  if (!response.ok) {
    throw new Error(`Hono ${method} ${path} returned ${response.status}: ${JSON.stringify(payload)}`)
  }

  return payload
}

function readSummaryPayload(payload: unknown, label: string): Summary {
  const record = readRecord(payload, `${label} response`)
  return readSummary(readRecord(record.summary, `${label} response.summary`))
}

function readEditPayload(payload: unknown): EditResult {
  const record = readRecord(payload, 'edit response')
  const checks = readRecord(record.checks, 'edit response.checks')
  return {
    records: readNumber(record.records, 'edit response.records'),
    after: readSummary(readRecord(record.after, 'edit response.after')),
    checks: {
      totalRevenueChanged: readBoolean(checks.totalRevenueChanged, 'edit response.checks.totalRevenueChanged'),
      formulasPersisted: readBoolean(checks.formulasPersisted, 'edit response.checks.formulasPersisted'),
      serializedBytes: readNumber(checks.serializedBytes, 'edit response.checks.serializedBytes'),
    },
  }
}

function readSummary(record: Record<string, unknown>): Summary {
  return {
    totalRevenue: readNumber(record.totalRevenue, 'summary.totalRevenue'),
    westCustomers: readNumber(record.westCustomers, 'summary.westCustomers'),
    largestDeal: readNumber(record.largestDeal, 'summary.largestDeal'),
  }
}

function readRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be a JSON object`)
  }
  return value
}

function isJsonRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`${label} must be a number`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

function assertProof(result: HonoRouteProof): void {
  const expectedBefore: Summary = {
    totalRevenue: 36900,
    westCustomers: 20,
    largestDeal: 24000,
  }
  const expectedAfter: Summary = {
    totalRevenue: 48600,
    westCustomers: 20,
    largestDeal: 24000,
  }
  const failures: string[] = []

  if (!sameSummary(result.before, expectedBefore)) {
    failures.push(`before summary was ${JSON.stringify(result.before)}`)
  }
  if (!sameSummary(result.edit.after, expectedAfter)) {
    failures.push(`edit summary was ${JSON.stringify(result.edit.after)}`)
  }
  if (!sameSummary(result.after, expectedAfter)) {
    failures.push(`final summary was ${JSON.stringify(result.after)}`)
  }
  if (result.edit.records !== 4) {
    failures.push(`edited ${result.edit.records} records instead of 4`)
  }
  if (!result.edit.checks.totalRevenueChanged) {
    failures.push('totalRevenueChanged was false')
  }
  if (!result.edit.checks.formulasPersisted) {
    failures.push('formulasPersisted was false')
  }
  if (result.edit.checks.serializedBytes <= 0) {
    failures.push('serializedBytes was not positive')
  }

  if (failures.length > 0) {
    throw new Error(`Hono route smoke failed:\n${failures.map((failure) => `- ${failure}`).join('\n')}`)
  }
}

function sameSummary(left: Summary, right: Summary): boolean {
  return left.totalRevenue === right.totalRevenue && left.westCustomers === right.westCustomers && left.largestDeal === right.largestDeal
}
