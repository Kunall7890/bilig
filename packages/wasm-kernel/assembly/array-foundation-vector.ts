import { ValueTag } from './protocol'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { allocateOutputString, encodeOutputStringId, writeOutputStringData, writeSpillArrayValue } from './vm'

export function writeLiteralTextToSpill(arrayIndex: u32, offset: i32, text: string): void {
  const outputIndex = allocateOutputString(text.length)
  for (let index = 0; index < text.length; index += 1) {
    writeOutputStringData(outputIndex, index, <u16>text.charCodeAt(index))
  }
  writeSpillArrayValue(arrayIndex, offset, <u8>ValueTag.String, encodeOutputStringId(outputIndex))
}

export function vectorLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts)
  if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
    return i32.MIN_VALUE
  }
  return rows == 1 || cols == 1 ? rows * cols : i32.MIN_VALUE
}

export function vectorTagAt(
  slot: i32,
  rows: i32,
  cols: i32,
  index: i32,
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
  const row = rows == 1 ? 0 : index
  const col = rows == 1 ? index : 0
  return inputCellTag(
    slot,
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
}

export function vectorValueAt(
  slot: i32,
  rows: i32,
  cols: i32,
  index: i32,
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
  const row = rows == 1 ? 0 : index
  const col = rows == 1 ? index : 0
  return inputCellScalarValue(
    slot,
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

export function writeVectorValueToSpill(
  arrayIndex: u32,
  offset: i32,
  slot: i32,
  rows: i32,
  cols: i32,
  index: i32,
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
): void {
  writeSpillArrayValue(
    arrayIndex,
    offset,
    vectorTagAt(
      slot,
      rows,
      cols,
      index,
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
    vectorValueAt(
      slot,
      rows,
      cols,
      index,
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

export function findBucketIndex(tags: Array<u8>, values: Array<f64>, tag: u8, value: f64): i32 {
  for (let index = 0; index < tags.length; index += 1) {
    if (unchecked(tags[index]) == tag && unchecked(values[index]) == value) {
      return index
    }
  }
  return -1
}
