import { createInMemoryWorkbookStorage, createWorkPaperRequestHandler, type WorkPaperJsonStorage } from './route.ts'

type Summary = {
  totalRevenue: number
  westCustomers: number
  largestDeal: number
}
type SummaryResponse = {
  summary: Summary
}
type EditResponse = {
  records: number
  after: Summary
  checks: {
    totalRevenueChanged: boolean
    formulasPersisted: boolean
    serializedBytes: number
  }
}

const storageEvents: string[] = []
const inMemoryStorage = createInMemoryWorkbookStorage()
const storage: WorkPaperJsonStorage = {
  async loadWorkbookJson() {
    storageEvents.push('load')
    return inMemoryStorage.loadWorkbookJson()
  },
  async saveWorkbookJson(workbookJson: string) {
    storageEvents.push('save')
    await inMemoryStorage.saveWorkbookJson(workbookJson)
  },
}
const handleWorkPaperRequest = createWorkPaperRequestHandler(storage)

const updateRecords = [
  { region: 'West', customers: 20, arpa: 1200 },
  { region: 'East', customers: 30, arpa: 250 },
  { region: 'Central', customers: 18, arpa: 300 },
  { region: 'North', customers: 65, arpa: 180 },
]

const before = await requestJson('/api/workpaper/summary', parseSummaryResponse)
const edit = await requestJson('/api/workpaper/revenue', parseEditResponse, {
  method: 'POST',
  headers: {
    'content-type': 'application/json',
  },
  body: JSON.stringify({ records: updateRecords }),
})
const after = await requestJson('/api/workpaper/summary', parseSummaryResponse)

const output = {
  before: before.summary,
  edit: {
    records: edit.records,
    after: edit.after,
    checks: edit.checks,
  },
  after: after.summary,
  verified: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

async function requestJson<T>(path: string, parse: (value: unknown) => T, init?: RequestInit): Promise<T> {
  const response = await handleWorkPaperRequest(new Request(`http://localhost:8787${path}`, init))
  const body: unknown = await response.json()
  if (!response.ok) {
    throw new Error(`request failed: ${response.status} ${JSON.stringify(body)}`)
  }
  return parse(body)
}

function parseSummaryResponse(value: unknown): SummaryResponse {
  const record = readJsonRecord(value, 'summary response')
  return {
    summary: readSummary(record.summary, 'summary response summary'),
  }
}

function parseEditResponse(value: unknown): EditResponse {
  const record = readJsonRecord(value, 'edit response')
  const checks = readJsonRecord(record.checks, 'edit response checks')
  return {
    records: readNumber(record.records, 'edit response records'),
    after: readSummary(record.after, 'edit response after'),
    checks: {
      totalRevenueChanged: readBoolean(checks.totalRevenueChanged, 'edit response totalRevenueChanged'),
      formulasPersisted: readBoolean(checks.formulasPersisted, 'edit response formulasPersisted'),
      serializedBytes: readNumber(checks.serializedBytes, 'edit response serializedBytes'),
    },
  }
}

function readSummary(value: unknown, label: string): Summary {
  const record = readJsonRecord(value, label)
  return {
    totalRevenue: readNumber(record.totalRevenue, `${label} totalRevenue`),
    westCustomers: readNumber(record.westCustomers, `${label} westCustomers`),
    largestDeal: readNumber(record.largestDeal, `${label} largestDeal`),
  }
}

function readJsonRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isJsonRecord(value)) {
    throw new Error(`${label} must be an object`)
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

function assertOutput(actual: typeof output): void {
  const expectedBefore = {
    totalRevenue: 36900,
    westCustomers: 20,
    largestDeal: 24000,
  }
  const expectedAfter = {
    totalRevenue: 48600,
    westCustomers: 20,
    largestDeal: 24000,
  }

  if (
    JSON.stringify(actual.before) !== JSON.stringify(expectedBefore) ||
    JSON.stringify(actual.edit.after) !== JSON.stringify(expectedAfter) ||
    JSON.stringify(actual.after) !== JSON.stringify(expectedAfter) ||
    actual.edit.records !== 4 ||
    !actual.edit.checks.totalRevenueChanged ||
    !actual.edit.checks.formulasPersisted ||
    actual.edit.checks.serializedBytes <= 0 ||
    !storageEvents.includes('save')
  ) {
    throw new Error(`unexpected WorkPaper API result: ${JSON.stringify(actual)}`)
  }
}
