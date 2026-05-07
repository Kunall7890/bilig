import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
import { copyInputCellToSpill } from './array-materialize'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult } from './vm'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

export function tryApplyArrayFilterBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): i32 {
  if (builtinId != BuiltinId.Filter || (argc != 2 && argc != 3)) {
    return -1
  }

  const sourceKind = kindStack[base]
  const includeKind = kindStack[base + 1]
  if (
    (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) ||
    (includeKind != STACK_KIND_RANGE && includeKind != STACK_KIND_ARRAY)
  ) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
  const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
  const includeRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
  const includeCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
  if (
    sourceRows <= 0 ||
    sourceCols <= 0 ||
    includeRows <= 0 ||
    includeCols <= 0 ||
    sourceRows == i32.MIN_VALUE ||
    sourceCols == i32.MIN_VALUE ||
    includeRows == i32.MIN_VALUE ||
    includeCols == i32.MIN_VALUE
  ) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (includeRows == sourceRows && includeCols == 1) {
    const keptRows = new Array<i32>()
    for (let row = 0; row < sourceRows; row++) {
      const includeTag = inputCellTag(
        base + 1,
        row,
        0,
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
      const includeValue = inputCellScalarValue(
        base + 1,
        row,
        0,
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
      if (includeTag == ValueTag.Error) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, includeValue, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const keep = coerceBoolean(includeTag, includeValue)
      if (keep < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (keep == 1) {
        keptRows.push(row)
      }
    }
    if (keptRows.length == 0) {
      if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[base + 2],
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const arrayIndex = allocateSpillArrayResult(keptRows.length, sourceCols)
    let outputOffset = 0
    for (let index = 0; index < keptRows.length; index++) {
      const sourceRow = unchecked(keptRows[index])
      for (let col = 0; col < sourceCols; col++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          sourceRow,
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
        if (copyError != ErrorCode.None) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, keptRows.length, sourceCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (includeRows == 1 && includeCols == sourceCols) {
    const keptCols = new Array<i32>()
    for (let col = 0; col < sourceCols; col++) {
      const includeTag = inputCellTag(
        base + 1,
        0,
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
      const includeValue = inputCellScalarValue(
        base + 1,
        0,
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
      if (includeTag == ValueTag.Error) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, includeValue, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const keep = coerceBoolean(includeTag, includeValue)
      if (keep < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (keep == 1) {
        keptCols.push(col)
      }
    }
    if (keptCols.length == 0) {
      if (argc < 3 || kindStack[base + 2] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[base + 2],
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const arrayIndex = allocateSpillArrayResult(sourceRows, keptCols.length)
    let outputOffset = 0
    for (let row = 0; row < sourceRows; row++) {
      for (let index = 0; index < keptCols.length; index++) {
        const sourceCol = unchecked(keptCols[index])
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          row,
          sourceCol,
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
        if (copyError != ErrorCode.None) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, copyError, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        outputOffset += 1
      }
    }
    return writeArrayResult(base, arrayIndex, sourceRows, keptCols.length, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
}
