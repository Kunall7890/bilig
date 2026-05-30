import { ValueTag, type CellValue } from '@bilig/protocol'
import { parseNumericText } from '../numeric-text.js'
import { coerceScalarMathNumber } from './numeric.js'

export function toNumber(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
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

export function toAverageNumber(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

export function toScalarMathNumber(value: CellValue | undefined): number | undefined {
  return coerceScalarMathNumber(value, toNumber)
}

export function parseDirectAggregateNumericText(value: string): number | undefined {
  return parseNumericText(value)
}

export function toDirectAggregateNumber(value: CellValue): number | undefined {
  if (value.tag === ValueTag.String) {
    return parseDirectAggregateNumericText(value.value)
  }
  return toNumber(value)
}
