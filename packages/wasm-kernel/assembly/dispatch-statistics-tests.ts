import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { isNumericResult } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
import { chiSquareTestPValue, fTestPValue, tTestPValue, zTestPValue } from './statistics-tests'

function writeStatisticalTestResult(
  base: i32,
  result: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (result < 0.0) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, -result, rangeIndexStack, valueStack, tagStack, kindStack)
  }
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
    isNumericResult(result) ? result : ErrorCode.Value,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  )
}

export function tryApplyStatisticalTestBuiltin(
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
  if ((builtinId == BuiltinId.ChisqTest || builtinId == BuiltinId.Chitest || builtinId == BuiltinId.LegacyChitest) && argc == 2) {
    const result = chiSquareTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    )
    return writeStatisticalTestResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Ftest || builtinId == BuiltinId.FTest) && argc == 2) {
    const result = fTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    )
    return writeStatisticalTestResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Ztest || builtinId == BuiltinId.ZTest) && (argc == 2 || argc == 3)) {
    const result = zTestPValue(
      base,
      argc,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    )
    return writeStatisticalTestResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.TTest || builtinId == BuiltinId.Ttest) && argc == 4) {
    const result = tTestPValue(
      base,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    )
    return writeStatisticalTestResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
