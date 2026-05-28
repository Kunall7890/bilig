import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { parseArithmeticNumericText, parseNumericText } from './numeric-text.js'

export const numericReferenceAggregateCallees = new Set([
  'SUM',
  'AVERAGE',
  'AVG',
  'MIN',
  'MAX',
  'PRODUCT',
  'SUMSQ',
  'GCD',
  'LCM',
  'GEOMEAN',
  'HARMEAN',
])

export const numericReferenceStatisticCallees = new Set([
  'MODE',
  'MODE.SNGL',
  'STDEV',
  'STDEV.S',
  'STDEVP',
  'STDEV.P',
  'VAR',
  'VAR.S',
  'VARP',
  'VAR.P',
  'SKEW',
  'SKEW.P',
  'SKEWP',
  'KURT',
])

export const aStyleReferenceCallees = new Set(['AVERAGEA', 'MINA', 'MAXA'])
export const logicalReferenceCallees = new Set(['AND', 'OR', 'XOR'])

export function coerceDirectNumericTextAggregateArgument(callee: string, value: CellValue, isReferenceArg: boolean): CellValue {
  if (isReferenceArg || value.tag !== ValueTag.String) {
    return value
  }
  const directNumericText = value.value === '' ? 0 : parseNumericText(value.value)
  if (callee === 'COUNT') {
    return directNumericText === undefined ? value : { tag: ValueTag.Number, value: directNumericText }
  }
  if (callee === 'SUM' || callee === 'AVERAGE' || callee === 'AVG') {
    const numeric = parseArithmeticNumericText(value.value)
    return numeric === undefined ? { tag: ValueTag.Error, code: ErrorCode.Value } : { tag: ValueTag.Number, value: numeric }
  }
  if (callee === 'PRODUCT' || callee === 'MIN' || callee === 'MAX' || callee === 'SUMSQ') {
    return directNumericText === undefined
      ? { tag: ValueTag.Error, code: ErrorCode.Value }
      : { tag: ValueTag.Number, value: directNumericText }
  }
  if (callee === 'GEOMEAN' || callee === 'HARMEAN') {
    return directNumericText === undefined
      ? { tag: ValueTag.Error, code: ErrorCode.Value }
      : { tag: ValueTag.Number, value: directNumericText }
  }
  return value
}

export function coerceDirectNumericTextStatisticArgument(callee: string, value: CellValue, isReferenceArg: boolean): CellValue {
  if (isReferenceArg || value.tag !== ValueTag.String || !numericReferenceStatisticCallees.has(callee)) {
    return value
  }
  const numeric = value.value === '' ? 0 : parseNumericText(value.value)
  return numeric === undefined ? { tag: ValueTag.Error, code: ErrorCode.Value } : { tag: ValueTag.Number, value: numeric }
}

export function referencedScalarBuiltinAggregateValues(callee: string, value: CellValue): readonly CellValue[] | undefined {
  if (callee === 'COUNT') {
    return value.tag === ValueTag.Number ? [value] : []
  }
  if (!numericReferenceAggregateCallees.has(callee)) {
    return undefined
  }
  return value.tag === ValueTag.Number || value.tag === ValueTag.Error ? [value] : []
}

export function referencedScalarBuiltinStatisticValues(callee: string, value: CellValue): readonly CellValue[] | undefined {
  return numericReferenceStatisticCallees.has(callee) && (value.tag === ValueTag.Number || value.tag === ValueTag.Error)
    ? [value]
    : undefined
}

export function referencedScalarBuiltinLogicalValues(callee: string, value: CellValue): readonly CellValue[] | undefined {
  if (!logicalReferenceCallees.has(callee)) {
    return undefined
  }
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Error ? [value] : []
}

export function referencedScalarBuiltinAStyleValues(callee: string, value: CellValue): readonly CellValue[] | undefined {
  if (!aStyleReferenceCallees.has(callee)) {
    return undefined
  }
  return value.tag === ValueTag.Empty ? [] : [value]
}
