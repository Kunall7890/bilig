import { ValueTag, formatErrorCode, formatGeneralNumberValue, type CellValue } from '@bilig/protocol'
import { parseNumericText } from '@bilig/formula'

export function inlineRequiredNumber(value: CellValue): number | undefined {
  return value.tag === ValueTag.Error ? undefined : inlineNumber(value)
}

export function inlineOptionalNumber(value: CellValue | undefined, fallback: number): number | undefined {
  return value === undefined ? fallback : inlineRequiredNumber(value)
}

export function inlinePaymentType(value: CellValue | undefined, fallback: number): number | undefined {
  const numeric = inlineOptionalNumber(value, fallback)
  if (numeric === undefined) {
    return undefined
  }
  const type = Math.trunc(numeric)
  return type === 0 || type === 1 ? type : undefined
}

export function roundInlineToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits
    return Math.round(value * factor) / factor
  }
  const factor = 10 ** -digits
  return Math.round(value / factor) * factor
}

export function inlineNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

export function inlineArithmeticNumber(value: CellValue): number | undefined {
  if (value.tag !== ValueTag.String) {
    return inlineNumber(value)
  }
  const trimmed = value.value.trim()
  return trimmed === '' ? 0 : parseNumericText(trimmed)
}

export function inlineTruthy(value: CellValue): boolean {
  return (inlineNumber(value) ?? 0) !== 0
}

export function inlineStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return formatGeneralNumberValue(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return formatErrorCode(value.code)
  }
}

export function compareInlineScalars(left: CellValue, right: CellValue): number | undefined {
  if (left.tag === ValueTag.String && right.tag === ValueTag.String) {
    return compareInlineText(left.value, right.value)
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.Empty) {
    return 0
  }
  if (left.tag === ValueTag.String && right.tag === ValueTag.Empty) {
    return compareInlineText(left.value, '')
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.String) {
    return compareInlineText('', right.value)
  }
  if (left.tag === ValueTag.String && isInlineNumberLike(right)) {
    return 1
  }
  if (isInlineNumberLike(left) && right.tag === ValueTag.String) {
    return -1
  }
  const leftNumber = inlineNumber(left)
  const rightNumber = inlineNumber(right)
  if (leftNumber === undefined || rightNumber === undefined) {
    return undefined
  }
  if (leftNumber === rightNumber || (Object.is(leftNumber, -0) && Object.is(rightNumber, 0))) {
    return 0
  }
  return leftNumber < rightNumber ? -1 : 1
}

function isInlineNumberLike(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean
}

function compareInlineText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}
