import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { copyInputCellToSpill, uniqueColKey, uniqueRowKey, uniqueScalarKey } from './array-materialize'
import { inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot } from './operands'
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

export function tryApplyArrayUniqueBuiltin(
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
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId != BuiltinId.Unique || argc < 1 || argc > 3) {
    return -1
  }

  const sourceKind = kindStack[base]
  if (sourceKind != STACK_KIND_RANGE && sourceKind != STACK_KIND_ARRAY) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (argc >= 2 && kindStack[base + 1] != STACK_KIND_SCALAR) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (argc >= 3 && kindStack[base + 2] != STACK_KIND_SCALAR) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (argc >= 2 && tagStack[base + 1] == ValueTag.Error) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base + 1], rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (argc >= 3 && tagStack[base + 2] == ValueTag.Error) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base + 2], rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const byColFlag = argc >= 2 ? coerceBoolean(tagStack[base + 1], valueStack[base + 1]) : 0
  const exactlyOnceFlag = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0
  if (byColFlag < 0 || exactlyOnceFlag < 0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
  const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
  if (sourceRows <= 0 || sourceCols <= 0 || sourceRows == i32.MIN_VALUE || sourceCols == i32.MIN_VALUE) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (sourceRows == 1 || sourceCols == 1) {
    const vectorLength = sourceRows * sourceCols
    const keys = new Array<string>()
    for (let index = 0; index < vectorLength; index++) {
      const row = sourceRows == 1 ? 0 : index
      const col = sourceRows == 1 ? index : 0
      const tag = inputCellTag(
        base,
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
      const value = inputCellScalarValue(
        base,
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
      if (tag == ValueTag.Error) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const key = uniqueScalarKey(
        tag,
        value,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (key == null) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      keys.push(key)
    }

    const keptIndexes = new Array<i32>()
    for (let index = 0; index < vectorLength; index++) {
      const key = unchecked(keys[index])
      let seenEarlier = false
      for (let prior = 0; prior < index; prior++) {
        if (unchecked(keys[prior]) == key) {
          seenEarlier = true
          break
        }
      }
      if (seenEarlier) {
        continue
      }
      if (exactlyOnceFlag == 1) {
        let count = 0
        for (let cursor = 0; cursor < vectorLength; cursor++) {
          if (unchecked(keys[cursor]) == key) {
            count += 1
          }
        }
        if (count != 1) {
          continue
        }
      }
      keptIndexes.push(index)
    }

    const outputRows = sourceRows == 1 ? 1 : keptIndexes.length
    const outputCols = sourceRows == 1 ? keptIndexes.length : 1
    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols)
    for (let outputOffset = 0; outputOffset < keptIndexes.length; outputOffset++) {
      const sourceIndex = unchecked(keptIndexes[outputOffset])
      const sourceRow = sourceRows == 1 ? 0 : sourceIndex
      const sourceCol = sourceRows == 1 ? sourceIndex : 0
      const copyError = copyInputCellToSpill(
        arrayIndex,
        outputOffset,
        base,
        sourceRow,
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
    }
    return writeArrayResult(base, arrayIndex, outputRows, outputCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (byColFlag == 1) {
    const keys = new Array<string>()
    for (let col = 0; col < sourceCols; col++) {
      const key = uniqueColKey(
        base,
        col,
        sourceRows,
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
      if (key == null) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      keys.push(key)
    }

    const keptCols = new Array<i32>()
    for (let col = 0; col < sourceCols; col++) {
      const key = unchecked(keys[col])
      let seenEarlier = false
      for (let prior = 0; prior < col; prior++) {
        if (unchecked(keys[prior]) == key) {
          seenEarlier = true
          break
        }
      }
      if (seenEarlier) {
        continue
      }
      if (exactlyOnceFlag == 1) {
        let count = 0
        for (let cursor = 0; cursor < sourceCols; cursor++) {
          if (unchecked(keys[cursor]) == key) {
            count += 1
          }
        }
        if (count != 1) {
          continue
        }
      }
      keptCols.push(col)
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

  const keys = new Array<string>()
  for (let row = 0; row < sourceRows; row++) {
    const key = uniqueRowKey(
      base,
      row,
      sourceCols,
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
    if (key == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    keys.push(key)
  }

  const keptRows = new Array<i32>()
  for (let row = 0; row < sourceRows; row++) {
    const key = unchecked(keys[row])
    let seenEarlier = false
    for (let prior = 0; prior < row; prior++) {
      if (unchecked(keys[prior]) == key) {
        seenEarlier = true
        break
      }
    }
    if (seenEarlier) {
      continue
    }
    if (exactlyOnceFlag == 1) {
      let count = 0
      for (let cursor = 0; cursor < sourceRows; cursor++) {
        if (unchecked(keys[cursor]) == key) {
          count += 1
        }
      }
      if (count != 1) {
        continue
      }
    }
    keptRows.push(row)
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
