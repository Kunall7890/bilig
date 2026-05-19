import { ErrorCode, ValueTag } from './protocol'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { allocateOutputString, encodeOutputStringId, scalarText, writeOutputStringData, writeSpillArrayValue } from './vm'
import { generalNumberText, trimAsciiWhitespace } from './text-codec'

export function writeLiteralTextToSpill(arrayIndex: u32, offset: i32, text: string): void {
  const outputIndex = allocateOutputString(text.length)
  for (let index = 0; index < text.length; index += 1) {
    writeOutputStringData(outputIndex, index, <u16>text.charCodeAt(index))
  }
  writeSpillArrayValue(arrayIndex, offset, <u8>ValueTag.String, encodeOutputStringId(outputIndex))
}

function errorText(code: i32): string {
  if (code == ErrorCode.Div0) return '#DIV/0!'
  if (code == ErrorCode.Ref) return '#REF!'
  if (code == ErrorCode.Value) return '#VALUE!'
  if (code == ErrorCode.Name) return '#NAME?'
  if (code == ErrorCode.NA) return '#N/A'
  if (code == ErrorCode.Cycle) return '#CYCLE!'
  if (code == ErrorCode.Spill) return '#SPILL!'
  if (code == ErrorCode.Blocked) return '#BLOCKED!'
  return '#ERROR!'
}

export function writeScalarTextToSpill(arrayIndex: u32, offset: i32, tag: u8, value: f64): void {
  if (tag == ValueTag.String) {
    writeSpillArrayValue(arrayIndex, offset, tag, value)
    return
  }
  if (tag == ValueTag.Empty) {
    writeLiteralTextToSpill(arrayIndex, offset, '')
    return
  }
  if (tag == ValueTag.Boolean) {
    writeLiteralTextToSpill(arrayIndex, offset, value != 0 ? 'TRUE' : 'FALSE')
    return
  }
  if (tag == ValueTag.Number) {
    const text = generalNumberText(value)
    writeLiteralTextToSpill(arrayIndex, offset, text == null ? '' : text)
    return
  }
  writeLiteralTextToSpill(arrayIndex, offset, errorText(<i32>value))
}

export function writeScalarTextWithFallbackToSpill(arrayIndex: u32, offset: i32, tag: u8, value: f64, fallbackText: string): void {
  if (scalarTextIsBlank(tag, value)) {
    writeLiteralTextToSpill(arrayIndex, offset, fallbackText)
    return
  }
  writeScalarTextToSpill(arrayIndex, offset, tag, value)
}

function scalarTextIsBlank(tag: u8, value: f64): bool {
  if (tag == ValueTag.Empty) {
    return true
  }
  if (tag == ValueTag.Number) {
    const text = generalNumberText(value)
    return text == null || trimAsciiWhitespace(text).length == 0
  }
  if (tag == ValueTag.Boolean || tag == ValueTag.Error) {
    return false
  }
  if (tag == ValueTag.String) {
    const text = scalarText(tag, value)
    return text == null || trimAsciiWhitespace(text).length == 0
  }
  return true
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

export function writeVectorTextToSpill(
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
  writeScalarTextToSpill(
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

export function writeVectorTextWithFallbackToSpill(
  arrayIndex: u32,
  offset: i32,
  slot: i32,
  rows: i32,
  cols: i32,
  index: i32,
  fallbackText: string,
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
  writeScalarTextWithFallbackToSpill(
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
    fallbackText,
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
