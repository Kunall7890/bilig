import * as XLSX from 'xlsx'

import type { LiteralInput } from '@bilig/protocol'
import { toLiteralInput } from './workbook-import-helpers.js'

const legacyExcelErrorTextByCode = new Map<number, string>([
  [0, '#NULL!'],
  [7, '#DIV/0!'],
  [15, '#VALUE!'],
  [23, '#REF!'],
  [29, '#NAME?'],
  [36, '#NUM!'],
  [42, '#N/A'],
  [43, '#GETTING_DATA'],
])

export function readImportedNumberFormat(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed === 'General') {
    return undefined
  }
  return trimmed
}

export function readImportedLiteralCellValue(cell: Record<string, unknown>): LiteralInput | undefined {
  if (cell['t'] === 'e') {
    if (Object.hasOwn(cell, 'w')) {
      const displayText = toLiteralInput(cell['w'])
      if (typeof displayText === 'string' && displayText.startsWith('#')) {
        return displayText
      }
    }
    if (!Object.hasOwn(cell, 'v')) {
      return undefined
    }
    const errorCode = cell['v']
    if (errorCode === undefined || errorCode === null) {
      return undefined
    }
    if (typeof errorCode === 'number') {
      return legacyExcelErrorTextByCode.get(errorCode) ?? '#ERROR!'
    }
    if (typeof errorCode === 'string' && errorCode.startsWith('#')) {
      return errorCode
    }
    return '#ERROR!'
  }
  return toLiteralInput(cell['v'])
}

export function compareCellAddresses(left: string, right: string): number {
  const leftCell = XLSX.utils.decode_cell(left)
  const rightCell = XLSX.utils.decode_cell(right)
  return leftCell.r - rightCell.r || leftCell.c - rightCell.c
}
