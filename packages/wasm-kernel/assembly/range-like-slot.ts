import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { toNumberOrNaN, toNumberOrZero } from './operands'
import { ValueTag } from './protocol'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE } from './result-io'

export function isRangeLikeSlot(kind: u8): bool {
  return kind == STACK_KIND_RANGE || kind == STACK_KIND_ARRAY
}

export function rangeLikeSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (!isRangeLikeSlot(kindStack[slot])) {
    return i32.MIN_VALUE
  }
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return rows > 0 && cols > 0 ? rows * cols : i32.MIN_VALUE
}

function rangeLikeSlotRow(slot: i32, offset: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeColCounts: Uint32Array): i32 {
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return cols > 0 ? offset / cols : i32.MIN_VALUE
}

function rangeLikeSlotCol(slot: i32, offset: i32, kindStack: Uint8Array, rangeIndexStack: Uint32Array, rangeColCounts: Uint32Array): i32 {
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  return cols > 0 ? offset % cols : i32.MIN_VALUE
}

export function rangeLikeSlotTagAt(
  slot: i32,
  offset: i32,
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
): u8 {
  return inputCellTag(
    slot,
    rangeLikeSlotRow(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
    rangeLikeSlotCol(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
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
}

export function rangeLikeSlotValueAt(
  slot: i32,
  offset: i32,
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
): f64 {
  return inputCellScalarValue(
    slot,
    rangeLikeSlotRow(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
    rangeLikeSlotCol(slot, offset, kindStack, rangeIndexStack, rangeColCounts),
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

export function rangeLikeSlotNumberOrZero(
  slot: i32,
  offset: i32,
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
): f64 {
  return toNumberOrZero(
    rangeLikeSlotTagAt(
      slot,
      offset,
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
    ),
    rangeLikeSlotValueAt(
      slot,
      offset,
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
    ),
  )
}

export function rangeLikeSlotNumberOrNaN(
  slot: i32,
  offset: i32,
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
): f64 {
  return toNumberOrNaN(
    rangeLikeSlotTagAt(
      slot,
      offset,
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
    ),
    rangeLikeSlotValueAt(
      slot,
      offset,
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
    ),
  )
}

export function rangeLikeSlotNumberOnlyOrNaN(
  slot: i32,
  offset: i32,
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
): f64 {
  const tag = rangeLikeSlotTagAt(
    slot,
    offset,
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
  if (tag != ValueTag.Number) {
    return NaN
  }
  return rangeLikeSlotValueAt(
    slot,
    offset,
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
