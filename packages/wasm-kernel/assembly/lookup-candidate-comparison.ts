import { ValueTag } from './protocol'
import { scalarText } from './text-codec'
import { inputCellScalarValue, inputCellTag } from './operands'
import { slotVectorCol, slotVectorRow } from './lookup-slot'

function lookupNumber(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return NaN
}

function sameExactLookupNumber(left: f64, right: f64): bool {
  if (left == right) {
    return true
  }
  if (!isFinite(left) || !isFinite(right)) {
    return false
  }
  const leftMagnitude = Math.abs(left)
  const rightMagnitude = Math.abs(right)
  const scale = Math.max(1, Math.max(leftMagnitude, rightMagnitude))
  return Math.abs(left - right) <= scale * 5e-15
}

export function compareLookupScalars(
  leftTag: u8,
  leftValue: f64,
  rightTag: u8,
  rightValue: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const leftTextlike = leftTag == ValueTag.String || leftTag == ValueTag.Empty
  const rightTextlike = rightTag == ValueTag.String || rightTag == ValueTag.Empty
  if (leftTextlike && rightTextlike) {
    const leftText = scalarText(
      leftTag,
      leftValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const rightText = scalarText(
      rightTag,
      rightValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (leftText == null || rightText == null) {
      return i32.MIN_VALUE
    }
    const normalizedLeft = leftText.toUpperCase()
    const normalizedRight = rightText.toUpperCase()
    if (normalizedLeft == normalizedRight) {
      return 0
    }
    return normalizedLeft < normalizedRight ? -1 : 1
  }

  const leftNumber = lookupNumber(leftTag, leftValue)
  const rightNumber = lookupNumber(rightTag, rightValue)
  if (isNaN(leftNumber) || isNaN(rightNumber)) {
    return i32.MIN_VALUE
  }
  if (sameExactLookupNumber(leftNumber, rightNumber)) {
    return 0
  }
  return leftNumber < rightNumber ? -1 : 1
}

export function compareLookupCell(
  sourceSlot: i32,
  row: i32,
  col: i32,
  lookupTag: u8,
  lookupValue: f64,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const candidateTag = inputCellTag(
    sourceSlot,
    row,
    col,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
  )
  const candidateValue = inputCellScalarValue(
    sourceSlot,
    row,
    col,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  )
  return compareLookupScalars(
    candidateTag,
    candidateValue,
    lookupTag,
    lookupValue,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}

export function compareLookupVectorCandidate(
  sourceSlot: i32,
  index: i32,
  lookupTag: u8,
  lookupValue: f64,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  return compareLookupCell(
    sourceSlot,
    slotVectorRow(sourceSlot, index, kindStack, rangeIndexStack, rangeRowCounts),
    slotVectorCol(sourceSlot, index, kindStack, rangeIndexStack, rangeRowCounts),
    lookupTag,
    lookupValue,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}
