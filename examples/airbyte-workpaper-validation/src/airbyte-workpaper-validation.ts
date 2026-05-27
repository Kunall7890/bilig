import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

export interface AirbyteRecordMessage {
  readonly type: 'RECORD'
  readonly record: {
    readonly stream: string
    readonly data: {
      readonly order_id: string
      readonly customer_id: string
      readonly updated_at: string
      readonly amount: number
      readonly status: 'paid' | 'rejected'
    }
    readonly emitted_at: number
  }
}

export interface AirbyteStateMessage {
  readonly type: 'STATE'
  readonly state: {
    readonly state_type?: 'STREAM' | 'LEGACY' | 'GLOBAL'
    readonly stream?: {
      readonly stream_descriptor?: {
        readonly name?: string
      }
      readonly stream_state?: {
        readonly cursor?: string
      }
    }
    readonly data?: {
      readonly cursor?: string
    }
    readonly global?: {
      readonly shared_state?: {
        readonly cursor?: string
      }
      readonly stream_states?: readonly {
        readonly stream_descriptor?: {
          readonly name?: string
        }
        readonly stream_state?: {
          readonly cursor?: string
        }
      }[]
    }
  }
}

export type AirbyteMessage = AirbyteRecordMessage | AirbyteStateMessage

export interface AirbyteWorkPaperValidationInput {
  readonly initialStateCursor?: string
  readonly expectedPaidAmount: number
  readonly expectedRecordCount: number
  readonly messages: readonly AirbyteMessage[]
  readonly stream?: string
}

export interface AirbyteWorkPaperValidationResult {
  readonly patch: {
    readonly stream: string
    readonly state_type: 'STREAM' | 'LEGACY' | 'GLOBAL'
    readonly committed_state_cursor: string
    readonly record_count: number
    readonly gross_amount: number
    readonly paid_amount: number
    readonly rejected_records: number
    readonly validation_passed: boolean
  }
  readonly proof: {
    readonly editedCells: readonly ['Inputs!B2', 'Inputs!B3', 'Inputs!B4', 'Inputs!B5']
    readonly stateCursorSource: string
    readonly before: AirbyteWorkPaperValidationSummary
    readonly after: AirbyteWorkPaperValidationSummary
    readonly afterRestore: AirbyteWorkPaperValidationSummary
    readonly persistedDocumentBytes: number
    readonly verified: boolean
  }
  readonly limitations: readonly string[]
}

export interface AirbyteWorkPaperValidationSummary {
  readonly recordCount: number
  readonly grossAmount: number
  readonly paidAmount: number
  readonly rejectedRecords: number
  readonly lastRecordCursor: string
  readonly lastRecordCursorMillis: number
  readonly committedStateCursor: string
  readonly committedStateCursorMillis: number
  readonly stateCursorMatchesRecords: boolean
  readonly expectedPaidAmount: number
  readonly paidAmountMatchesExpected: boolean
  readonly expectedRecordCount: number
  readonly recordCountMatchesExpected: boolean
}

interface NormalizedOrderRecord {
  readonly orderId: string
  readonly customerId: string
  readonly cursor: string
  readonly cursorMillis: number
  readonly amount: number
  readonly status: 'paid' | 'rejected'
  readonly paidFlag: 0 | 1
  readonly rejectedFlag: 0 | 1
}

interface FinalStateCursor {
  readonly cursor: string
  readonly stateType: 'STREAM' | 'LEGACY' | 'GLOBAL'
  readonly source: string
}

export function readAirbyteMessagesFromJsonl(jsonl: string): AirbyteMessage[] {
  return jsonl
    .split(/\r?\n/u)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .map((line, index) => parseAirbyteMessage(line, index + 1))
}

export function validateAirbyteOrdersWithWorkPaper(input: AirbyteWorkPaperValidationInput): AirbyteWorkPaperValidationResult {
  const stream = input.stream ?? 'orders'
  const records = input.messages.filter(
    (message): message is AirbyteRecordMessage => message.type === 'RECORD' && message.record.stream === stream,
  )
  const normalizedRecords = records.map(normalizeOrderRecord)
  const finalState = readFinalStateCursor(input.messages, stream)
  const expectedPaidAmount = readFiniteNumber(input.expectedPaidAmount, 'expectedPaidAmount')
  const expectedRecordCount = readInteger(input.expectedRecordCount, 'expectedRecordCount')
  const initialStateCursor = input.initialStateCursor ?? finalState.cursor
  const initialStateCursorMillis = readCursorMillis(initialStateCursor, 'initialStateCursor')
  const finalStateCursorMillis = readCursorMillis(finalState.cursor, 'finalStateCursor')
  const lastRecord = normalizedRecords.at(-1)

  if (lastRecord === undefined) {
    throw new Error(`Airbyte stream ${stream} did not include any RECORD messages`)
  }

  const workbook = WorkPaper.buildFromSheets({
    Inputs: [
      ['Field', 'Value'],
      ['Committed state cursor', initialStateCursor],
      ['Committed state cursor millis', initialStateCursorMillis],
      ['Expected paid amount', expectedPaidAmount],
      ['Expected record count', expectedRecordCount],
    ],
    Records: buildRecordsSheet(normalizedRecords),
    Summary: [
      ['Metric', 'Value'],
      ['Record count', buildSumFormula(normalizedRecords.map(() => '1'))],
      ['Gross amount', buildSheetSumFormula('Records', 'E', normalizedRecords.length)],
      ['Paid amount', buildSheetSumFormula('Records', 'H', normalizedRecords.length)],
      ['Rejected records', buildSheetSumFormula('Records', 'I', normalizedRecords.length)],
      ['Last record cursor', lastRecord.cursor],
      ['Last record cursor millis', lastRecord.cursorMillis],
      ['Committed state cursor', '=Inputs!B2'],
      ['Committed state cursor millis', '=Inputs!B3'],
      ['State cursor matches records', '=B7=B9'],
      ['Expected paid amount', '=Inputs!B4'],
      ['Paid amount matches expected', '=B4=B11'],
      ['Expected record count', '=Inputs!B5'],
      ['Record count matches expected', '=B2=B13'],
    ],
  })

  try {
    const inputsSheet = requireSheet(workbook, 'Inputs')
    const summarySheet = requireSheet(workbook, 'Summary')
    const before = readSummary(workbook, summarySheet)

    workbook.setCellContents({ sheet: inputsSheet, row: 1, col: 1 }, finalState.cursor)
    workbook.setCellContents({ sheet: inputsSheet, row: 2, col: 1 }, finalStateCursorMillis)
    workbook.setCellContents({ sheet: inputsSheet, row: 3, col: 1 }, expectedPaidAmount)
    workbook.setCellContents({ sheet: inputsSheet, row: 4, col: 1 }, expectedRecordCount)
    const after = readSummary(workbook, summarySheet)

    const document = exportWorkPaperDocument(workbook, { includeConfig: true })
    const serialized = serializeWorkPaperDocument(document)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))

    try {
      const restoredSummarySheet = requireSheet(restored, 'Summary')
      const afterRestore = readSummary(restored, restoredSummarySheet)
      const verified =
        sameSummary(after, afterRestore) &&
        after.stateCursorMatchesRecords &&
        after.paidAmountMatchesExpected &&
        after.recordCountMatchesExpected &&
        serialized.length > 0

      return {
        patch: {
          stream,
          state_type: finalState.stateType,
          committed_state_cursor: after.committedStateCursor,
          record_count: after.recordCount,
          gross_amount: after.grossAmount,
          paid_amount: after.paidAmount,
          rejected_records: after.rejectedRecords,
          validation_passed: verified,
        },
        proof: {
          editedCells: ['Inputs!B2', 'Inputs!B3', 'Inputs!B4', 'Inputs!B5'],
          stateCursorSource: finalState.source,
          before,
          after,
          afterRestore,
          persistedDocumentBytes: new TextEncoder().encode(serialized).byteLength,
          verified,
        },
        limitations: [
          'This example validates an Airbyte-style record/state export after sync; it is not an Airbyte connector or official Airbyte integration.',
          'Airbyte state is source-defined; keep source, destination, and platform checkpoint semantics as the authority for replication correctness.',
          'Keep warehouse constraints, destination acknowledgements, Airbyte job metadata, and domain data-quality checks in the loop for production pipelines.',
        ],
      }
    } finally {
      restored.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function parseAirbyteMessage(line: string, lineNumber: number): AirbyteMessage {
  let parsed: unknown
  try {
    parsed = JSON.parse(line) as unknown
  } catch (error) {
    throw new Error(`Invalid Airbyte JSONL on line ${String(lineNumber)}: ${String(error)}`, { cause: error })
  }

  if (!isAirbyteMessage(parsed)) {
    throw new Error(`Unsupported Airbyte message on line ${String(lineNumber)}: ${line}`)
  }

  return parsed
}

function isAirbyteMessage(value: unknown): value is AirbyteMessage {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }

  const type = Reflect.get(value, 'type')
  if (type === 'RECORD') {
    const record = Reflect.get(value, 'record')
    return isObject(record) && typeof Reflect.get(record, 'stream') === 'string' && isObject(Reflect.get(record, 'data'))
  }

  if (type === 'STATE') {
    return isObject(Reflect.get(value, 'state'))
  }

  return false
}

function normalizeOrderRecord(message: AirbyteRecordMessage): NormalizedOrderRecord {
  const data = message.record.data
  const status = data.status

  return {
    orderId: readNonEmptyString(data.order_id, 'order_id'),
    customerId: readNonEmptyString(data.customer_id, 'customer_id'),
    cursor: readNonEmptyString(data.updated_at, 'updated_at'),
    cursorMillis: readCursorMillis(data.updated_at, 'updated_at'),
    amount: readFiniteNumber(data.amount, 'amount'),
    status,
    paidFlag: status === 'paid' ? 1 : 0,
    rejectedFlag: status === 'rejected' ? 1 : 0,
  }
}

function readFinalStateCursor(messages: readonly AirbyteMessage[], stream: string): FinalStateCursor {
  const stateMessages = messages.filter((message): message is AirbyteStateMessage => message.type === 'STATE')
  for (let index = stateMessages.length - 1; index >= 0; index -= 1) {
    const state = stateMessages[index]?.state

    const globalCursor = readGlobalStateCursor(state?.global, stream)
    if (globalCursor !== undefined) {
      return globalCursor
    }

    const streamName = state?.stream?.stream_descriptor?.name
    const streamCursor = state?.stream?.stream_state?.cursor
    if ((streamName === undefined || streamName === stream) && streamCursor !== undefined) {
      return {
        cursor: readNonEmptyString(streamCursor, 'state.stream.stream_state.cursor'),
        stateType: 'STREAM',
        source: 'state.stream.stream_state.cursor',
      }
    }

    const legacyCursor = state?.data?.cursor
    if (legacyCursor !== undefined) {
      return {
        cursor: readNonEmptyString(legacyCursor, 'state.data.cursor'),
        stateType: 'LEGACY',
        source: 'state.data.cursor',
      }
    }
  }

  throw new Error(`Airbyte stream ${stream} did not include a usable STATE cursor`)
}

function readGlobalStateCursor(
  globalState: AirbyteStateMessage['state']['global'] | undefined,
  stream: string,
): FinalStateCursor | undefined {
  if (globalState === undefined) {
    return undefined
  }

  for (const streamState of globalState.stream_states ?? []) {
    if (streamState.stream_descriptor?.name !== stream) {
      continue
    }

    const streamCursor = streamState.stream_state?.cursor
    if (streamCursor !== undefined) {
      return {
        cursor: readNonEmptyString(streamCursor, 'state.global.stream_states[].stream_state.cursor'),
        stateType: 'GLOBAL',
        source: 'state.global.stream_states[].stream_state.cursor',
      }
    }
  }

  const sharedCursor = globalState.shared_state?.cursor
  if (sharedCursor !== undefined) {
    return {
      cursor: readNonEmptyString(sharedCursor, 'state.global.shared_state.cursor'),
      stateType: 'GLOBAL',
      source: 'state.global.shared_state.cursor',
    }
  }

  return undefined
}

function buildRecordsSheet(records: readonly NormalizedOrderRecord[]): (readonly (number | string)[])[] {
  return [
    ['order_id', 'customer_id', 'cursor', 'cursor_millis', 'amount', 'status', 'paid_flag', 'paid_amount', 'rejected_flag'],
    ...records.map((record, index) => {
      const row = index + 2
      return [
        record.orderId,
        record.customerId,
        record.cursor,
        record.cursorMillis,
        record.amount,
        record.status,
        record.paidFlag,
        `=E${row}*G${row}`,
        record.rejectedFlag,
      ]
    }),
  ]
}

function buildSheetSumFormula(sheetName: string, column: string, rowCount: number): string {
  return buildSumFormula(Array.from({ length: rowCount }, (_, index) => `${sheetName}!${column}${String(index + 2)}`))
}

function buildSumFormula(terms: readonly string[]): string {
  if (terms.length === 0) {
    return '=0'
  }

  return `=${terms.join('+')}`
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)

  if (sheetId === undefined) {
    throw new Error(`Missing ${sheetName} sheet`)
  }

  return sheetId
}

function readSummary(workbook: WorkPaper, summarySheet: number): AirbyteWorkPaperValidationSummary {
  return {
    recordCount: readNumberCell(workbook, summarySheet, 1, 'record count'),
    grossAmount: readNumberCell(workbook, summarySheet, 2, 'gross amount'),
    paidAmount: readNumberCell(workbook, summarySheet, 3, 'paid amount'),
    rejectedRecords: readNumberCell(workbook, summarySheet, 4, 'rejected records'),
    lastRecordCursor: readStringCell(workbook, summarySheet, 5, 'last record cursor'),
    lastRecordCursorMillis: readNumberCell(workbook, summarySheet, 6, 'last record cursor millis'),
    committedStateCursor: readStringCell(workbook, summarySheet, 7, 'committed state cursor'),
    committedStateCursorMillis: readNumberCell(workbook, summarySheet, 8, 'committed state cursor millis'),
    stateCursorMatchesRecords: readBooleanCell(workbook, summarySheet, 9, 'state cursor matches records'),
    expectedPaidAmount: readNumberCell(workbook, summarySheet, 10, 'expected paid amount'),
    paidAmountMatchesExpected: readBooleanCell(workbook, summarySheet, 11, 'paid amount matches expected'),
    expectedRecordCount: readNumberCell(workbook, summarySheet, 12, 'expected record count'),
    recordCountMatchesExpected: readBooleanCell(workbook, summarySheet, 13, 'record count matches expected'),
  }
}

function readCellValue(workbook: WorkPaper, sheet: number, row: number, col: number): unknown {
  const cell = workbook.getCellValue({ sheet, row, col })
  return typeof cell === 'object' && cell !== null ? Reflect.get(cell, 'value') : cell
}

function readNumberCell(workbook: WorkPaper, sheet: number, row: number, label: string): number {
  return readFiniteNumber(readCellValue(workbook, sheet, row, 1), label)
}

function readStringCell(workbook: WorkPaper, sheet: number, row: number, label: string): string {
  return readNonEmptyString(readCellValue(workbook, sheet, row, 1), label)
}

function readBooleanCell(workbook: WorkPaper, sheet: number, row: number, label: string): boolean {
  const value = readCellValue(workbook, sheet, row, 1)

  if (typeof value !== 'boolean') {
    throw new Error(`Expected boolean ${label}, got ${JSON.stringify(value)}`)
  }

  return value
}

function readCursorMillis(value: string, label: string): number {
  const cursor = readNonEmptyString(value, label)
  const millis = Date.parse(cursor)

  if (!Number.isFinite(millis)) {
    throw new Error(`${label} must be an ISO timestamp cursor`)
  }

  return millis
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

function isObject(value: unknown): value is object {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function sameSummary(left: AirbyteWorkPaperValidationSummary, right: AirbyteWorkPaperValidationSummary): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
