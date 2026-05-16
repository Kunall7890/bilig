import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { inputCellNumeric, inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot, toNumberOrNaN } from './operands'
import { coerceInteger } from './numeric-core'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

export function tryApplyStatisticsRankBuiltin(
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
  if (!((builtinId == BuiltinId.Rank || builtinId == BuiltinId.RankEq || builtinId == BuiltinId.RankAvg) && (argc == 2 || argc == 3))) {
    return -1
  }

  if (kindStack[base] != STACK_KIND_SCALAR) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  if (tagStack[base] == ValueTag.Error) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const target = toNumberOrNaN(tagStack[base], valueStack[base])
  if (!isFinite(target)) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  let order = 0
  if (argc == 3) {
    if (kindStack[base + 2] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const requestedOrder = coerceInteger(tagStack[base + 2], valueStack[base + 2])
    if (requestedOrder == i32.MIN_VALUE || (requestedOrder != 0 && requestedOrder != 1)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    order = requestedOrder
  }

  const arraySlot = base + 1
  const rows = inputRowsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeRowCounts)
  const cols = inputColsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeColCounts)
  if (rows < 1 || cols < 1) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  let preceding = 0
  let ties = 0
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const valueTag = inputCellTag(
        arraySlot,
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
      if (valueTag == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          inputCellScalarValue(
            arraySlot,
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
          ),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }

      const numeric = inputCellNumeric(
        arraySlot,
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
      if (!isFinite(numeric)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }

      if (numeric == target) {
        ties += 1
        continue
      }
      if ((order == 0 && numeric > target) || (order == 1 && numeric < target)) {
        preceding += 1
      }
    }
  }

  if (ties == 0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const rank = builtinId == BuiltinId.RankAvg ? <f64>preceding + (<f64>ties + 1.0) / 2.0 : <f64>(preceding + 1)
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, rank, rangeIndexStack, valueStack, tagStack, kindStack)
}
