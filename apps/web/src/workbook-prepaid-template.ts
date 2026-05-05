import type { CommitOp } from '@bilig/core'
import {
  buildCellNumberFormatCode,
  type CellNumberFormatInput,
  type CellRangeRef,
  type CellStylePatch,
  type LiteralInput,
} from '@bilig/protocol'
import type { WorkbookMutationMethod } from './workbook-sync.js'

export interface WorkbookTemplateMutation {
  readonly method: WorkbookMutationMethod
  readonly args: readonly unknown[]
}

interface PrepaidExample {
  readonly vendor: string
  readonly description: string
  readonly start: readonly [year: number, month: number, day: number]
  readonly end: readonly [year: number, month: number, day: number]
  readonly amount: number
  readonly notes: string
}

const TEMPLATE_RANGE_START = 'A1'
const TEMPLATE_RANGE_END = 'W40'
const TEMPLATE_HEADER_ROW = 5
const TEMPLATE_FIRST_DATA_ROW = 6
const TEMPLATE_LAST_DATA_ROW = 10
const TEMPLATE_YEAR = 2024
const EXCEL_EPOCH_OFFSET = 25569
const DAY_MS = 86_400_000

const PREPAID_EXAMPLES: readonly PrepaidExample[] = [
  {
    vendor: 'ABC Insurance',
    description: 'General Liability',
    start: [2024, 4, 1],
    end: [2025, 3, 31],
    amount: 12_000,
    notes: 'Annual policy',
  },
  {
    vendor: 'Northstar SaaS',
    description: 'Revenue platform subscription',
    start: [2024, 1, 15],
    end: [2025, 1, 14],
    amount: 24_000,
    notes: 'Prorated first and final month',
  },
  {
    vendor: 'Metro Properties',
    description: 'Office rent prepayment',
    start: [2024, 2, 1],
    end: [2024, 7, 31],
    amount: 18_000,
    notes: 'Six-month lease incentive',
  },
  {
    vendor: 'Compliance Cloud',
    description: 'Audit workflow license',
    start: [2024, 5, 10],
    end: [2024, 11, 9],
    amount: 7_500,
    notes: 'Daily amortization',
  },
  {
    vendor: 'Cybersecurity Policy',
    description: 'Cyber insurance',
    start: [2024, 10, 1],
    end: [2025, 9, 30],
    amount: 9_600,
    notes: 'Open balance carries forward',
  },
] as const

const HEADERS = [
  'Vendor',
  'Description',
  'Start Date',
  'End Date',
  'Total Amount',
  'Life Months',
  'Monthly Average',
  'Jan 2024',
  'Feb 2024',
  'Mar 2024',
  'Apr 2024',
  'May 2024',
  'Jun 2024',
  'Jul 2024',
  'Aug 2024',
  'Sep 2024',
  'Oct 2024',
  'Nov 2024',
  'Dec 2024',
  '2024 Amortized',
  'Remaining Balance',
  'Status',
  'Notes',
] as const

const COLUMN_WIDTHS = [184, 190, 104, 104, 118, 94, 118, 92, 92, 92, 92, 92, 92, 92, 92, 92, 92, 92, 92, 130, 134, 118, 180] as const

const CURRENCY_FORMAT: CellNumberFormatInput = {
  kind: 'currency',
  currency: 'USD',
  decimals: 2,
  useGrouping: true,
  negativeStyle: 'minus',
  zeroStyle: 'zero',
}

const INTEGER_FORMAT: CellNumberFormatInput = {
  kind: 'number',
  decimals: 0,
  useGrouping: true,
}

const ISO_DATE_FORMAT: CellNumberFormatInput = {
  kind: 'date',
  dateStyle: 'iso',
}

const CURRENCY_FORMAT_CODE = buildCellNumberFormatCode(CURRENCY_FORMAT)
const INTEGER_FORMAT_CODE = buildCellNumberFormatCode(INTEGER_FORMAT)
const ISO_DATE_FORMAT_CODE = buildCellNumberFormatCode(ISO_DATE_FORMAT)

const THIN_BORDER = {
  color: '#DDD8CC',
  style: 'solid',
  weight: 'thin',
} as const

function range(sheetName: string, startAddress: string, endAddress = startAddress): CellRangeRef {
  return { sheetName, startAddress, endAddress }
}

function columnName(columnIndex: number): string {
  let index = columnIndex + 1
  let name = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    index = Math.floor((index - 1) / 26)
  }
  return name
}

function excelDateSerial(year: number, month: number, day: number): number {
  return Date.UTC(year, month - 1, day) / DAY_MS + EXCEL_EPOCH_OFFSET
}

function monthFormula(row: number, month: number): string {
  return `ROUND(IFERROR($E${row}*MAX(0,MIN($D${row},EOMONTH(DATE(${TEMPLATE_YEAR},${month},1),0))-MAX($C${row},DATE(${TEMPLATE_YEAR},${month},1))+1)/($D${row}-$C${row}+1),0),2)`
}

function pushCell(ops: CommitOp[], sheetName: string, addr: string, value: LiteralInput, format?: string): void {
  ops.push({ kind: 'upsertCell', sheetName, addr, value, ...(format ? { format } : {}) })
}

function pushFormula(ops: CommitOp[], sheetName: string, addr: string, formula: string, format?: string): void {
  ops.push({ kind: 'upsertCell', sheetName, addr, formula, ...(format ? { format } : {}) })
}

function pushStyle(
  mutations: WorkbookTemplateMutation[],
  sheetName: string,
  startAddress: string,
  endAddress: string,
  patch: CellStylePatch,
): void {
  mutations.push({ method: 'setRangeStyle', args: [range(sheetName, startAddress, endAddress), patch] })
}

function pushNumberFormat(
  mutations: WorkbookTemplateMutation[],
  sheetName: string,
  startAddress: string,
  endAddress: string,
  format: CellNumberFormatInput,
): void {
  mutations.push({ method: 'setRangeNumberFormat', args: [range(sheetName, startAddress, endAddress), format] })
}

function buildPrepaidTemplateCells(sheetName: string): CommitOp[] {
  const ops: CommitOp[] = []
  pushCell(ops, sheetName, 'A1', 'Prepaid Amortization Schedule')
  pushCell(ops, sheetName, 'A2', 'Daily-prorated 2024 amortization with summary totals, status, and open balance tracking.')
  pushCell(ops, sheetName, 'A3', 'Prepaids')
  pushFormula(ops, sheetName, 'B3', `ROUND(SUM(E${TEMPLATE_FIRST_DATA_ROW}:E${TEMPLATE_LAST_DATA_ROW}),2)`, CURRENCY_FORMAT_CODE)
  pushCell(ops, sheetName, 'C3', 'Avg/mo')
  pushFormula(ops, sheetName, 'D3', `ROUND(SUM(G${TEMPLATE_FIRST_DATA_ROW}:G${TEMPLATE_LAST_DATA_ROW}),2)`, CURRENCY_FORMAT_CODE)
  pushCell(ops, sheetName, 'E3', 'Amortized')
  pushFormula(ops, sheetName, 'F3', `ROUND(SUM(T${TEMPLATE_FIRST_DATA_ROW}:T${TEMPLATE_LAST_DATA_ROW}),2)`, CURRENCY_FORMAT_CODE)
  pushCell(ops, sheetName, 'G3', 'Remaining')
  pushFormula(ops, sheetName, 'H3', `ROUND(SUM(U${TEMPLATE_FIRST_DATA_ROW}:U${TEMPLATE_LAST_DATA_ROW}),2)`, CURRENCY_FORMAT_CODE)
  pushCell(ops, sheetName, 'A4', 'Inputs A:E | Schedule H:S | Totals T:U | Status V')

  HEADERS.forEach((header, columnIndex) => {
    pushCell(ops, sheetName, `${columnName(columnIndex)}${TEMPLATE_HEADER_ROW}`, header)
  })

  PREPAID_EXAMPLES.forEach((example, exampleIndex) => {
    const row = TEMPLATE_FIRST_DATA_ROW + exampleIndex
    pushCell(ops, sheetName, `A${row}`, example.vendor)
    pushCell(ops, sheetName, `B${row}`, example.description)
    pushCell(ops, sheetName, `C${row}`, excelDateSerial(...example.start), ISO_DATE_FORMAT_CODE)
    pushCell(ops, sheetName, `D${row}`, excelDateSerial(...example.end), ISO_DATE_FORMAT_CODE)
    pushCell(ops, sheetName, `E${row}`, example.amount, CURRENCY_FORMAT_CODE)
    pushFormula(ops, sheetName, `F${row}`, `IFERROR(DATEDIF(C${row},D${row}+1,"M"),0)`, INTEGER_FORMAT_CODE)
    pushFormula(ops, sheetName, `G${row}`, `ROUND(IFERROR(E${row}/F${row},0),2)`, CURRENCY_FORMAT_CODE)
    for (let month = 1; month <= 12; month += 1) {
      pushFormula(ops, sheetName, `${columnName(6 + month)}${row}`, monthFormula(row, month), CURRENCY_FORMAT_CODE)
    }
    pushFormula(ops, sheetName, `T${row}`, `ROUND(SUM(H${row}:S${row}),2)`, CURRENCY_FORMAT_CODE)
    pushFormula(ops, sheetName, `U${row}`, `ROUND(E${row}-T${row},2)`, CURRENCY_FORMAT_CODE)
    pushFormula(ops, sheetName, `V${row}`, `IF(U${row}<=0,"Complete",IF(T${row}=0,"Not started","In progress"))`)
    pushCell(ops, sheetName, `W${row}`, example.notes)
  })

  return ops
}

export function buildPrepaidAmortizationTemplateMutations(sheetName: string): WorkbookTemplateMutation[] {
  const mutations: WorkbookTemplateMutation[] = [
    { method: 'unmergeCells', args: [range(sheetName, TEMPLATE_RANGE_START, TEMPLATE_RANGE_END)] },
    { method: 'clearRange', args: [range(sheetName, TEMPLATE_RANGE_START, TEMPLATE_RANGE_END)] },
    { method: 'renderCommit', args: [buildPrepaidTemplateCells(sheetName)] },
    { method: 'mergeCells', args: [range(sheetName, 'A1', 'W1')] },
    { method: 'mergeCells', args: [range(sheetName, 'A2', 'W2')] },
    { method: 'mergeCells', args: [range(sheetName, 'A4', 'W4')] },
  ]

  pushStyle(mutations, sheetName, 'A1', 'W1', {
    fill: { backgroundColor: '#21563A' },
    font: { bold: true, color: '#FFFFFF', size: 14 },
    alignment: { horizontal: 'left', vertical: 'middle' },
  })
  pushStyle(mutations, sheetName, 'A2', 'W2', {
    fill: { backgroundColor: '#E8F1EB' },
    font: { color: '#21563A', size: 11 },
    alignment: { vertical: 'middle' },
  })
  pushStyle(mutations, sheetName, 'A5', 'W5', {
    fill: { backgroundColor: '#21563A' },
    font: { bold: true, color: '#FFFFFF', size: 11 },
    alignment: { horizontal: 'center', vertical: 'middle' },
    borders: { bottom: { color: '#173D29', style: 'solid', weight: 'medium' } },
  })
  pushStyle(mutations, sheetName, 'A3', 'H4', {
    fill: { backgroundColor: '#F7F7F4' },
    font: { color: '#1F2933', size: 11 },
    alignment: { vertical: 'middle' },
    borders: { bottom: THIN_BORDER },
  })
  pushStyle(mutations, sheetName, 'A3', 'G3', {
    font: { bold: true, color: '#52606D', size: 10 },
  })
  pushStyle(mutations, sheetName, 'B3', 'H3', {
    font: { bold: true, color: '#1F2933', size: 12 },
  })
  pushStyle(mutations, sheetName, 'A6', 'W10', {
    font: { color: '#1F2933', size: 11 },
    alignment: { vertical: 'middle' },
    borders: { bottom: THIN_BORDER },
  })
  pushStyle(mutations, sheetName, 'H6', 'S10', {
    fill: { backgroundColor: '#F4F8F5' },
  })
  pushStyle(mutations, sheetName, 'T6', 'U10', {
    fill: { backgroundColor: '#FFF7E6' },
    font: { bold: true },
  })
  pushStyle(mutations, sheetName, 'V6', 'V10', {
    fill: { backgroundColor: '#F3F2EE' },
    alignment: { horizontal: 'center', vertical: 'middle' },
    font: { bold: true, color: '#52606D' },
  })

  pushNumberFormat(mutations, sheetName, 'B3', 'B3', CURRENCY_FORMAT)
  pushNumberFormat(mutations, sheetName, 'D3', 'D3', CURRENCY_FORMAT)
  pushNumberFormat(mutations, sheetName, 'F3', 'F3', CURRENCY_FORMAT)
  pushNumberFormat(mutations, sheetName, 'H3', 'H3', CURRENCY_FORMAT)
  pushNumberFormat(mutations, sheetName, 'C6', 'D10', ISO_DATE_FORMAT)
  pushNumberFormat(mutations, sheetName, 'E6', 'E10', CURRENCY_FORMAT)
  pushNumberFormat(mutations, sheetName, 'F6', 'F10', INTEGER_FORMAT)
  pushNumberFormat(mutations, sheetName, 'G6', 'U10', CURRENCY_FORMAT)

  COLUMN_WIDTHS.forEach((width, columnIndex) => {
    mutations.push({ method: 'updateColumnMetadata', args: [sheetName, columnIndex, 1, width, null] })
  })
  mutations.push({ method: 'setFreezePane', args: [sheetName, TEMPLATE_HEADER_ROW, 0] })

  return mutations
}
