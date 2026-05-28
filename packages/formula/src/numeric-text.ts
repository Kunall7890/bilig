import type { ExcelDateSystem } from './builtins/excel-date.js'
import { parseDateValueFromText, parseElapsedTimeValueText } from './date-time-text.js'

const simpleNumericTextPattern = /^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+-]?\d+)?$/

function trimAsciiWhitespace(value: string): string {
  let start = 0
  let end = value.length
  while (start < end && value.charCodeAt(start) <= 32) {
    start += 1
  }
  while (end > start && value.charCodeAt(end - 1) <= 32) {
    end -= 1
  }
  return value.slice(start, end)
}

export function parseNumericText(value: string): number | undefined {
  const trimmed = trimAsciiWhitespace(value)
  if (trimmed.length === 0) {
    return undefined
  }

  if (simpleNumericTextPattern.test(trimmed)) {
    const direct = Number(trimmed)
    if (Number.isFinite(direct)) {
      return direct
    }
  }

  if (!trimmed.includes(',')) {
    return undefined
  }

  const grouped = /^([+-]?)(\d{1,3}(?:,\d{3})+)(\.\d*)?([eE][+-]?\d+)?$/.exec(trimmed)
  if (!grouped) {
    return undefined
  }

  const normalized = `${grouped[1] ?? ''}${(grouped[2] ?? '').replaceAll(',', '')}${grouped[3] ?? ''}${grouped[4] ?? ''}`
  const numeric = Number(normalized)
  return Number.isFinite(numeric) ? numeric : undefined
}

export function parseArithmeticNumericText(value: string): number | undefined {
  return value === '' ? 0 : parseNumericText(value)
}

export function parseArithmeticScalarText(value: string, dateSystem: ExcelDateSystem = '1900'): number | undefined {
  if (value === '') {
    return 0
  }
  const numeric = parseNumericText(value)
  if (numeric !== undefined) {
    return numeric
  }
  const date = parseDateValueFromText(value, dateSystem)
  const time = parseElapsedTimeValueText(value)
  if (date !== undefined) {
    return date + (time ?? 0)
  }
  return time
}
