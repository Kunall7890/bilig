import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'

export interface NormalizedCommandRange {
  readonly range: CellRangeRef
  readonly cellCount: number
}

export interface CommandRangeBounds {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly startRow: number
  readonly startCol: number
  readonly endRow: number
  readonly endCol: number
}

const commandRangeDataFields = Object.freeze(['sheetName', 'startAddress', 'endAddress'] as const)

export function normalizeCommandRange(value: unknown, path: string, label = 'Workbook command bundle'): NormalizedCommandRange {
  if (!isRecord(value)) {
    throw new Error(`${label} ${path} must be an object`)
  }
  for (const field of commandRangeDataFields) {
    const descriptor = Object.getOwnPropertyDescriptor(value, field)
    if (descriptor !== undefined && !('value' in descriptor)) {
      throw new Error(`${label} ${path}.${field} must be a data property`)
    }
  }
  const sheetName = ownDataValue(value, 'sheetName')
  const startAddress = ownDataValue(value, 'startAddress')
  const endAddress = ownDataValue(value, 'endAddress')
  if (typeof sheetName !== 'string' || typeof startAddress !== 'string' || typeof endAddress !== 'string') {
    throw new Error(`${label} ${path} must include sheetName, startAddress, and endAddress strings`)
  }
  const normalizedSheetName = normalizeExactString(sheetName, `${path}.sheetName`, label)
  const start = normalizeCellAddress(startAddress, `${path}.startAddress`, label)
  const end = normalizeCellAddress(endAddress, `${path}.endAddress`, label)
  if (end.row < start.row || end.col < start.col) {
    throw new Error(`${label} ${path} endAddress must not be before startAddress`)
  }
  return {
    range: Object.freeze({
      sheetName: normalizedSheetName,
      startAddress: start.text,
      endAddress: end.text,
    }),
    cellCount: (end.row - start.row + 1) * (end.col - start.col + 1),
  }
}

export function commandRangeBounds(value: CellRangeRef, path: string, label: string): CommandRangeBounds {
  const normalized = normalizeCommandRange(value, path, label).range
  const start = parseCellAddress(normalized.startAddress)
  const end = parseCellAddress(normalized.endAddress)
  return {
    sheetName: normalized.sheetName,
    startAddress: normalized.startAddress,
    endAddress: normalized.endAddress,
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  }
}

export function commandRangeContains(outer: CommandRangeBounds, inner: CommandRangeBounds): boolean {
  return (
    outer.sheetName === inner.sheetName &&
    inner.startRow >= outer.startRow &&
    inner.startCol >= outer.startCol &&
    inner.endRow <= outer.endRow &&
    inner.endCol <= outer.endCol
  )
}

export function commandRangeLabel(range: CommandRangeBounds): string {
  return range.startAddress === range.endAddress
    ? `${range.sheetName}!${range.startAddress}`
    : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

function normalizeCellAddress(
  value: string,
  path: string,
  label: string,
): { readonly row: number; readonly col: number; readonly text: string } {
  try {
    const parsed = parseCellAddress(value)
    if (parsed.sheetName !== undefined) {
      throw new Error('qualified')
    }
    return {
      row: parsed.row,
      col: parsed.col,
      text: formatAddress(parsed.row, parsed.col),
    }
  } catch {
    throw new Error(`${label} ${path} is invalid: ${value}`)
  }
}

function normalizeExactString(value: string, path: string, label: string): string {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`${label} ${path} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`${label} ${path} must not have leading or trailing whitespace`)
  }
  return normalized
}

function ownDataValue(value: Record<string, unknown>, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}
