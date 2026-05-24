import type { CellRangeRef } from '@bilig/protocol'
import { normalizeCommandRange } from './command-ranges.js'

export interface WorkbookCommandReceiptRangeIssue {
  readonly path: string
  readonly message: string
}

const commandReceiptRangeDataFields = Object.freeze(['sheetName', 'startAddress', 'endAddress'] as const)

export function commandReceiptChangedRangeIssues(value: unknown): readonly WorkbookCommandReceiptRangeIssue[] {
  if (!Array.isArray(value)) {
    return Object.freeze([rangeIssue('changedRanges', 'Workbook command receipt changed ranges must be an array')])
  }

  const issues: WorkbookCommandReceiptRangeIssue[] = []
  for (let index = 0; index < value.length; index += 1) {
    const path = `changedRanges[${String(index)}]`
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(rangeIssue(path, 'Workbook command receipt changed ranges must contain only data properties'))
      continue
    }
    const dataFieldIssue = firstRangeDataFieldIssue(descriptor.value, path)
    if (dataFieldIssue !== null) {
      issues.push(dataFieldIssue)
      continue
    }
    try {
      normalizeCommandRange(descriptor.value, path, 'Workbook command receipt')
    } catch (error) {
      issues.push(rangeIssue(path, errorMessage(error)))
    }
  }
  return Object.freeze(issues)
}

function firstRangeDataFieldIssue(value: unknown, path: string): WorkbookCommandReceiptRangeIssue | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return null
  }
  for (const field of commandReceiptRangeDataFields) {
    const descriptor = Object.getOwnPropertyDescriptor(value, field)
    if (descriptor !== undefined && !('value' in descriptor)) {
      return rangeIssue(`${path}.${field}`, `Workbook command receipt ${path}.${field} must be a data property`)
    }
  }
  return null
}

export function normalizeCommandReceiptChangedRanges(value: unknown): readonly CellRangeRef[] | null {
  if (commandReceiptChangedRangeIssues(value).length > 0 || !Array.isArray(value)) {
    return null
  }

  const ranges: CellRangeRef[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      return null
    }
    ranges.push(normalizeCommandRange(descriptor.value, `changedRanges[${String(index)}]`, 'Workbook command receipt').range)
  }
  return Object.freeze(ranges)
}

function rangeIssue(path: string, message: string): WorkbookCommandReceiptRangeIssue {
  return Object.freeze({
    path,
    message,
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
