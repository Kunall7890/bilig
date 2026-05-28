import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname } from 'node:path'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

interface MeltanoWorkPaperInput {
  readonly expectedPaidAmount: number
  readonly expectedRecordCount: number
  readonly maxRejectedRecords: number
  readonly output: string
  readonly previousExpectedPaidAmount: number
  readonly previousExpectedRecordCount: number
  readonly records: string
}

interface OrderRecord {
  readonly amount: number
  readonly customerId: string
  readonly orderId: string
  readonly paidFlag: 0 | 1
  readonly rejectedFlag: 0 | 1
  readonly status: 'paid' | 'rejected'
  readonly updatedAt: string
}

interface MeltanoWorkPaperSummary {
  readonly expectedPaidAmount: number
  readonly expectedRecordCount: number
  readonly maxRejectedRecords: number
  readonly paidAmount: number
  readonly paidAmountMatchesExpected: boolean
  readonly recordCount: number
  readonly recordCountMatchesExpected: boolean
  readonly rejectedRecords: number
  readonly rejectedWithinMax: boolean
}

interface MeltanoWorkPaperResult {
  readonly patch: {
    readonly command: 'meltano invoke bilig-workpaper-validator:validate'
    readonly paid_amount: number
    readonly record_count: number
    readonly rejected_records: number
    readonly validation_passed: boolean
  }
  readonly proof: {
    readonly after: MeltanoWorkPaperSummary
    readonly afterRestore: MeltanoWorkPaperSummary
    readonly before: MeltanoWorkPaperSummary
    readonly editedCells: readonly ['Inputs!B2', 'Inputs!B4']
    readonly outputFile: string
    readonly persistedDocumentBytes: number
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

const cliInput = readCliInput(process.argv.slice(2))
const result = await validateMeltanoRecordsWithWorkPaper(cliInput)

await mkdir(dirname(cliInput.output), { recursive: true })
await writeFile(cliInput.output, `${JSON.stringify(result, null, 2)}\n`, 'utf8')

console.log(JSON.stringify(result, null, 2))

if (!result.proof.verified) {
  throw new Error(`Meltano WorkPaper validation proof failed: ${JSON.stringify(result.proof)}`)
}

async function validateMeltanoRecordsWithWorkPaper(input: MeltanoWorkPaperInput): Promise<MeltanoWorkPaperResult> {
  const records = parseRecords(await readFile(input.records, 'utf8'))
  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Expected record count', input.previousExpectedRecordCount],
      ['Max rejected records', input.maxRejectedRecords],
      ['Expected paid amount', input.previousExpectedPaidAmount],
      ['Output artifact', input.output],
    ],
    Records: buildRecordsSheet(records),
    Summary: [
      ['Metric', 'Value'],
      ['Record count', buildSumFormula(records.map(() => '1'))],
      ['Paid amount', buildSheetSumFormula('Records', 'F', records.length)],
      ['Rejected records', buildSheetSumFormula('Records', 'G', records.length)],
      ['Expected record count', '=Inputs!B2'],
      ['Record count matches expected', '=B2=B5'],
      ['Expected paid amount', '=Inputs!B4'],
      ['Paid amount matches expected', '=B3=B7'],
      ['Max rejected records', '=Inputs!B3'],
      ['Rejected within max', '=B4<=B9'],
    ],
  })

  try {
    const inputsSheet = requireSheet(workbook, 'Inputs')
    const summarySheet = requireSheet(workbook, 'Summary')
    const before = readSummary(workbook, summarySheet)

    workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, input.expectedRecordCount)
    workbook.setCellContents({ sheet: inputsSheet, row: 3, col: 1 }, input.expectedPaidAmount)
    const after = readSummary(workbook, summarySheet)

    const document = exportWorkPaperDocument(workbook, { includeConfig: true })
    const serialized = serializeWorkPaperDocument(document)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))

    try {
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const afterRestore = readSummary(restored, restoredSummarySheet)
      const verified =
        sameSummary(after, afterRestore) &&
        after.recordCountMatchesExpected &&
        after.paidAmountMatchesExpected &&
        after.rejectedWithinMax &&
        serialized.length > 0

      return {
        patch: {
          command: 'meltano invoke bilig-workpaper-validator:validate',
          paid_amount: after.paidAmount,
          record_count: after.recordCount,
          rejected_records: after.rejectedRecords,
          validation_passed: verified,
        },
        proof: {
          editedCells: ['Inputs!B2', 'Inputs!B4'],
          before,
          after,
          afterRestore,
          outputFile: input.output,
          persistedDocumentBytes: new TextEncoder().encode(serialized).byteLength,
          verified,
        },
        limitations: [
          'This example proves a Meltano utility command shape, not an accepted Meltano Hub listing.',
          'Keep Meltano run history, destination-table row counts, and loader acknowledgements as the pipeline authority.',
          'Use a warehouse query or destination export as the records input for production validation.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function parseRecords(jsonl: string): OrderRecord[] {
  return jsonl
    .split(/\r?\n/u)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .map((line, index) => parseRecord(line, index + 1))
}

function parseRecord(line: string, lineNumber: number): OrderRecord {
  let parsed: unknown
  try {
    parsed = JSON.parse(line) as unknown
  } catch (error) {
    throw new Error(`Invalid JSONL on line ${String(lineNumber)}: ${String(error)}`, { cause: error })
  }

  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`Record line ${String(lineNumber)} must be an object`)
  }

  const status = readStatus(Reflect.get(parsed, 'status'), `line ${String(lineNumber)} status`)
  const amount = readFiniteNumber(Reflect.get(parsed, 'amount'), `line ${String(lineNumber)} amount`)

  return {
    amount,
    customerId: readNonEmptyString(Reflect.get(parsed, 'customer_id'), `line ${String(lineNumber)} customer_id`),
    orderId: readNonEmptyString(Reflect.get(parsed, 'order_id'), `line ${String(lineNumber)} order_id`),
    paidFlag: status === 'paid' ? 1 : 0,
    rejectedFlag: status === 'rejected' ? 1 : 0,
    status,
    updatedAt: readNonEmptyString(Reflect.get(parsed, 'updated_at'), `line ${String(lineNumber)} updated_at`),
  }
}

function buildRecordsSheet(records: readonly OrderRecord[]): (readonly (number | string)[])[] {
  return [
    ['order_id', 'customer_id', 'amount', 'status', 'paid_flag', 'paid_amount', 'rejected_flag', 'updated_at'],
    ...records.map((record, index) => {
      const row = index + 2
      return [
        record.orderId,
        record.customerId,
        record.amount,
        record.status,
        record.paidFlag,
        `=C${row}*E${row}`,
        record.rejectedFlag,
        record.updatedAt,
      ]
    }),
  ]
}

function buildSheetSumFormula(sheetName: string, column: string, rowCount: number): string {
  return buildSumFormula(Array.from({ length: rowCount }, (_, index) => `${sheetName}!${column}${String(index + 2)}`))
}

function buildSumFormula(terms: readonly string[]): string {
  return terms.length === 0 ? '=0' : `=${terms.join('+')}`
}

function readSummary(workbook: WorkPaper, summarySheet: number): MeltanoWorkPaperSummary {
  return {
    recordCount: readNumberCell(workbook, summarySheet, 1, 'record count'),
    paidAmount: readNumberCell(workbook, summarySheet, 2, 'paid amount'),
    rejectedRecords: readNumberCell(workbook, summarySheet, 3, 'rejected records'),
    expectedRecordCount: readNumberCell(workbook, summarySheet, 4, 'expected record count'),
    recordCountMatchesExpected: readBooleanCell(workbook, summarySheet, 5, 'record count matches expected'),
    expectedPaidAmount: readNumberCell(workbook, summarySheet, 6, 'expected paid amount'),
    paidAmountMatchesExpected: readBooleanCell(workbook, summarySheet, 7, 'paid amount matches expected'),
    maxRejectedRecords: readNumberCell(workbook, summarySheet, 8, 'max rejected records'),
    rejectedWithinMax: readBooleanCell(workbook, summarySheet, 9, 'rejected within max'),
  }
}

function readCliInput(args: readonly string[]): MeltanoWorkPaperInput {
  const values = new Map<string, string>()

  for (let index = 0; index < args.length; index += 2) {
    const flag = args[index]
    const value = args[index + 1]
    if (flag === undefined || value === undefined || !flag.startsWith('--')) {
      throw new Error(`Expected --flag value arguments, got ${args.join(' ')}`)
    }
    values.set(flag, value)
  }

  const expectedPaidAmount = readFiniteNumber(Number(values.get('--expected-paid-amount')), '--expected-paid-amount')
  const expectedRecordCount = readInteger(Number(values.get('--expected-record-count')), '--expected-record-count')

  return {
    expectedPaidAmount,
    expectedRecordCount,
    maxRejectedRecords: readOptionalInteger(values.get('--max-rejected-records'), 1, '--max-rejected-records'),
    output: readNonEmptyString(values.get('--output'), '--output'),
    previousExpectedPaidAmount: readOptionalNumber(values.get('--previous-expected-paid-amount'), 250, '--previous-expected-paid-amount'),
    previousExpectedRecordCount: readOptionalInteger(values.get('--previous-expected-record-count'), 3, '--previous-expected-record-count'),
    records: readNonEmptyString(values.get('--records'), '--records'),
  }
}

function readCellValue(workbook: WorkPaper, sheet: number, row: number, col: number): unknown {
  const cell = workbook.getCellValue({ sheet, row, col })
  return typeof cell === 'object' && cell !== null ? Reflect.get(cell, 'value') : cell
}

function readNumberCell(workbook: WorkPaper, sheet: number, row: number, label: string): number {
  return readFiniteNumber(readCellValue(workbook, sheet, row, 1), label)
}

function readBooleanCell(workbook: WorkPaper, sheet: number, row: number, label: string): boolean {
  const value = readCellValue(workbook, sheet, row, 1)

  if (typeof value !== 'boolean') {
    throw new Error(`Expected boolean ${label}, got ${JSON.stringify(value)}`)
  }

  return value
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheet = workbook.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }
  return sheet
}

function readStatus(value: unknown, label: string): 'paid' | 'rejected' {
  if (value !== 'paid' && value !== 'rejected') {
    throw new Error(`${label} must be paid or rejected`)
  }
  return value
}

function readOptionalNumber(value: string | undefined, fallback: number, label: string): number {
  return value === undefined ? fallback : readFiniteNumber(Number(value), label)
}

function readOptionalInteger(value: string | undefined, fallback: number, label: string): number {
  return value === undefined ? fallback : readInteger(Number(value), label)
}

function readInteger(value: unknown, label: string): number {
  const number = readFiniteNumber(value, label)
  if (!Number.isInteger(number)) {
    throw new Error(`${label} must be an integer`)
  }
  return number
}

function readFiniteNumber(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${label} must be a finite number`)
  }
  return value
}

function readNonEmptyString(value: unknown, label: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`${label} must be a non-empty string`)
  }
  return value
}

function sameSummary(left: MeltanoWorkPaperSummary, right: MeltanoWorkPaperSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
