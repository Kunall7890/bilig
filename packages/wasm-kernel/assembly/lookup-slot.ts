import { ErrorCode, ValueTag } from './protocol'
import { copyInputCellToResult } from './array-materialize'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from './result-io'
import { writeSpillArrayValue } from './vm'

export function lookupVectorSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (kindStack[slot] == STACK_KIND_SCALAR) {
    return 1
  }
  if (kindStack[slot] != STACK_KIND_RANGE && kindStack[slot] != STACK_KIND_ARRAY) {
    return i32.MIN_VALUE
  }
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  if (rows <= 0 || cols <= 0 || (rows != 1 && cols != 1)) {
    return i32.MIN_VALUE
  }
  return kindStack[slot] == STACK_KIND_RANGE ? <i32>rangeLengths[rangeIndexStack[slot]] : rows * cols
}

export function slotVectorLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (kindStack[slot] != STACK_KIND_RANGE && kindStack[slot] != STACK_KIND_ARRAY) {
    return i32.MIN_VALUE
  }
  return lookupVectorSlotLength(slot, kindStack, rangeIndexStack, rangeLengths, rangeRowCounts, rangeColCounts)
}

export function slotVectorRow(
  slot: i32,
  index: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
): i32 {
  return inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts) == 1 ? 0 : index
}

export function slotVectorCol(
  slot: i32,
  index: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
): i32 {
  return inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts) == 1 ? index : 0
}

export function writeLookupInputCellResult(
  base: i32,
  sourceSlot: i32,
  row: i32,
  col: i32,
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
): i32 {
  const tag = inputCellTag(
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
  if (tag == ValueTag.Empty) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  return copyInputCellToResult(
    base,
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
}

export function writeLookupInputCellToSpill(
  arrayIndex: u32,
  outputOffset: i32,
  sourceSlot: i32,
  row: i32,
  col: i32,
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
): i32 {
  const tag = inputCellTag(
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
  if (tag == ValueTag.Empty) {
    writeSpillArrayValue(arrayIndex, outputOffset, <u8>ValueTag.Number, 0)
    return ErrorCode.None
  }
  const value = inputCellScalarValue(
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
  if (tag == ValueTag.Error && isNaN(value)) {
    return ErrorCode.Value
  }
  writeSpillArrayValue(arrayIndex, outputOffset, tag, value)
  return ErrorCode.None
}
