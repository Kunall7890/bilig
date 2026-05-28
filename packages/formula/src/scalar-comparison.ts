import { ValueTag, type CellValue } from '@bilig/protocol'
import { normalizeExactLookupNumber } from './builtins/lookup-core-helpers.js'

function isNumberLike(value: CellValue): boolean {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean
}

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function comparableNumber(value: CellValue): number | undefined {
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

export function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if (left.tag === ValueTag.String && right.tag === ValueTag.String) {
    return compareText(left.value, right.value)
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.Empty) {
    return 0
  }
  if (left.tag === ValueTag.String && right.tag === ValueTag.Empty) {
    return compareText(left.value, '')
  }
  if (left.tag === ValueTag.Empty && right.tag === ValueTag.String) {
    return compareText('', right.value)
  }
  if (left.tag === ValueTag.String && isNumberLike(right)) {
    return 1
  }
  if (isNumberLike(left) && right.tag === ValueTag.String) {
    return -1
  }

  const leftNumber = comparableNumber(left)
  const rightNumber = comparableNumber(right)
  if (leftNumber === undefined || rightNumber === undefined) {
    return undefined
  }
  const normalizedLeft = normalizeExactLookupNumber(leftNumber)
  const normalizedRight = normalizeExactLookupNumber(rightNumber)
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}
