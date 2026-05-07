import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { sameExactLookupNumber } from '@bilig/formula'
import type { RuntimeFormula } from '../runtime-state.js'
import type { RuntimeColumnSlice } from './runtime-column-store-service.js'

export function decodeErrorCode(rawCode: number | undefined): ErrorCode {
  return rawCode ?? ErrorCode.None
}

export function evaluationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

export function decodeRuntimeTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

export function cellValueCriteriaString(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(Object.is(value.value, -0) ? 0 : value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return String(value.code)
  }
}

export function referenceReplacementKey(sheetName: string, address: string): string {
  return `${sheetName.trim().toUpperCase()}!${address.trim().toUpperCase()}`
}

export function cellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return right.tag === ValueTag.Number && left.value === right.value
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code
  }
}

export function sameExactNumericValue(left: number, right: number): boolean {
  return sameExactLookupNumber(left, right)
}

export function directNumberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

export function directErrorResult(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

export function offsetDirectAggregateResult(directAggregate: RuntimeFormula['directAggregate'], value: CellValue): CellValue {
  const offset = directAggregate?.resultOffset
  return offset !== undefined && value.tag === ValueTag.Number ? directNumberResult(value.value + offset) : value
}

export function directCriteriaCacheValueKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${value.value}`
    case ValueTag.Error:
      return `r:${value.code}`
  }
}

export function numericLikeValueAt(slice: RuntimeColumnSlice, offset: number): number | undefined {
  const tag = decodeRuntimeTag(slice.tags[offset])
  switch (tag) {
    case ValueTag.Number:
      return slice.numbers[offset] ?? 0
    case ValueTag.Boolean:
      return (slice.numbers[offset] ?? 0) !== 0 ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

export function strictNumericAggregateCandidateAt(slice: RuntimeColumnSlice, offset: number): number | undefined {
  return slice.tags[offset] === ValueTag.Number ? (slice.numbers[offset] ?? 0) : undefined
}
