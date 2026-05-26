import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { tryApplyCriteriaConditionalBuiltin } from './dispatch-criteria-conditional'
import { scalarErrorAt, rangeErrorAt } from './builtin-args'
import { gcdPairCalc, lcmPairCalc } from './numeric-core'
import { toNumberOrNaN } from './operands'
import { STACK_KIND_ARRAY, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from './result-io'
import { textLength } from './text-codec'
import { coerceScalarNumberLikeText } from './text-special'
import {
  readCachedRangeSumTag,
  readCachedRangeSumValue,
  readSpillArrayLength,
  readSpillArrayNumber,
  readSpillArrayTag,
  writeCachedRangeSum,
} from './vm'

function writeAggregateError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

function isCountblankBlank(tag: u8, value: f64, stringLengths: Uint32Array, outputStringLengths: Uint32Array): bool {
  return tag == ValueTag.Empty || (tag == ValueTag.String && textLength(tag, value, stringLengths, outputStringLengths) == 0)
}

const EXCEL_INTEGER_LIMIT: f64 = 9007199254740992.0

function gcdLcmIntegerOrNaN(value: f64): f64 {
  if (!isFinite(value) || value < 0.0 || value >= EXCEL_INTEGER_LIMIT) {
    return NaN
  }
  return Math.floor(value)
}

function lcmPairOrNaN(left: f64, right: f64): f64 {
  const result = lcmPairCalc(left, right)
  return !isFinite(result) || result >= EXCEL_INTEGER_LIMIT ? NaN : result
}

function directScalarNumberLikeText(
  slot: i32,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  return coerceScalarNumberLikeText(
    tagStack[slot],
    valueStack[slot],
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
}

function referencedNumberOrNaN(tag: u8, value: f64): f64 {
  return tag == ValueTag.Number ? value : NaN
}

function spillReferencedNumberOrNaN(arrayIndex: u32, cursor: i32): f64 {
  return readSpillArrayTag(arrayIndex, cursor) == ValueTag.Number ? readSpillArrayNumber(arrayIndex, cursor) : NaN
}

export function tryApplyAggregateCriteriaBuiltin(
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
  if (builtinId == BuiltinId.Sum) {
    if (argc == 1 && kindStack[base] == STACK_KIND_RANGE) {
      const rangeIndex = rangeIndexStack[base]
      const cachedTag = readCachedRangeSumTag(rangeIndex)
      if (cachedTag != 0xff) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          cachedTag,
          readCachedRangeSumValue(rangeIndex),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }

      const start = rangeOffsets[rangeIndex]
      const length = <i32>rangeLengths[rangeIndex]
      let sum = 0.0
      for (let cursor = 0; cursor < length; cursor += 1) {
        const memberIndex = rangeMembers[start + cursor]
        const memberTag = cellTags[memberIndex]
        if (memberTag == ValueTag.Error) {
          const errorCode = cellErrors[memberIndex]
          writeCachedRangeSum(rangeIndex, <u8>ValueTag.Error, errorCode)
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        const numeric = referencedNumberOrNaN(memberTag, cellNumbers[memberIndex])
        if (!isNaN(numeric)) {
          sum += numeric
        }
      }

      writeCachedRangeSum(rangeIndex, <u8>ValueTag.Number, sum)
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = referencedNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric)) {
            sum += numeric
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = spillReferencedNumberOrNaN(arrayIndex, cursor)
          if (!isNaN(numeric)) {
            sum += numeric
          }
        }
        continue
      }
      const numeric = coerceScalarNumberLikeText(
        tagStack[slot],
        valueStack[slot],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (isNaN(numeric) && tagStack[slot] == ValueTag.String) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!isNaN(numeric)) {
        sum += numeric
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Avg) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let sum = 0.0
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = referencedNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric)) {
            sum += numeric
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = spillReferencedNumberOrNaN(arrayIndex, cursor)
          if (!isNaN(numeric)) {
            sum += numeric
            count += 1
          }
        }
        continue
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot])
      if (!isNaN(numeric)) {
        sum += numeric
        count += 1
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0.0 : sum / <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Min) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let minValue = Infinity
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = referencedNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric) && numeric < minValue) {
            minValue = numeric
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = spillReferencedNumberOrNaN(arrayIndex, cursor)
          if (!isNaN(numeric) && numeric < minValue) {
            minValue = numeric
            count += 1
          }
        }
        continue
      }
      const numeric = directScalarNumberLikeText(
        slot,
        valueStack,
        tagStack,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (isNaN(numeric) && tagStack[slot] == ValueTag.String) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!isNaN(numeric) && numeric < minValue) {
        minValue = numeric
        count += 1
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0.0 : minValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Max) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let maxValue = -Infinity
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = referencedNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (!isNaN(numeric) && numeric > maxValue) {
            maxValue = numeric
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = spillReferencedNumberOrNaN(arrayIndex, cursor)
          if (!isNaN(numeric) && numeric > maxValue) {
            maxValue = numeric
            count += 1
          }
        }
        continue
      }
      const numeric = directScalarNumberLikeText(
        slot,
        valueStack,
        tagStack,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (isNaN(numeric) && tagStack[slot] == ValueTag.String) {
        return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!isNaN(numeric) && numeric > maxValue) {
        maxValue = numeric
        count += 1
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0.0 : maxValue,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (cellTags[memberIndex] == ValueTag.Number) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          if (readSpillArrayTag(arrayIndex, cursor) == ValueTag.Number) {
            count += 1
          }
        }
        continue
      }
      if (
        !isNaN(
          directScalarNumberLikeText(
            slot,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          ),
        )
      ) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (cellTags[memberIndex] != ValueTag.Empty) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot])
        continue
      }
      if (tagStack[slot] != ValueTag.Empty) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Countblank) {
    let count = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          if (isCountblankBlank(cellTags[memberIndex], <f64>cellStringIds[memberIndex], stringLengths, outputStringLengths)) {
            count += 1
          }
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          if (
            isCountblankBlank(
              readSpillArrayTag(arrayIndex, cursor),
              readSpillArrayNumber(arrayIndex, cursor),
              stringLengths,
              outputStringLengths,
            )
          ) {
            count += 1
          }
        }
        continue
      }
      if (isCountblankBlank(tagStack[slot], valueStack[slot], stringLengths, outputStringLengths)) {
        count += 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    builtinId == BuiltinId.Gcd ||
    builtinId == BuiltinId.Lcm ||
    builtinId == BuiltinId.Product ||
    builtinId == BuiltinId.Geomean ||
    builtinId == BuiltinId.Harmean ||
    builtinId == BuiltinId.Sumsq
  ) {
    const scalarError = <i32>scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeAggregateError(base, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rangeError = <i32>(
      rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors)
    )
    if (rangeError >= 0) {
      return writeAggregateError(base, rangeError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let count = 0
    let product = 1.0
    let sumSquares = 0.0
    let gcdValue = 0.0
    let lcmValue = 0.0
    let logSum = 0.0
    let reciprocalSum = 0.0

    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const numeric = referencedNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex])
          if (isNaN(numeric)) {
            continue
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            if (numeric == 0.0) {
              return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            logSum += Math.log(numeric)
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            reciprocalSum += 1.0 / numeric
          } else if (builtinId == BuiltinId.Gcd) {
            const integer = gcdLcmIntegerOrNaN(numeric)
            if (isNaN(integer)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            gcdValue = count == 0 ? integer : gcdPairCalc(gcdValue, integer)
          } else if (builtinId == BuiltinId.Lcm) {
            const integer = gcdLcmIntegerOrNaN(numeric)
            if (isNaN(integer)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            lcmValue = count == 0 ? integer : lcmPairOrNaN(lcmValue, integer)
            if (isNaN(lcmValue)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric
          }
          count += 1
        }
        continue
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot]
        const length = readSpillArrayLength(arrayIndex)
        for (let cursor = 0; cursor < length; cursor += 1) {
          const numeric = spillReferencedNumberOrNaN(arrayIndex, cursor)
          if (isNaN(numeric)) {
            continue
          }
          if (builtinId == BuiltinId.Geomean) {
            if (numeric < 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            if (numeric == 0.0) {
              return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            logSum += Math.log(numeric)
          } else if (builtinId == BuiltinId.Harmean) {
            if (numeric <= 0.0) {
              return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            reciprocalSum += 1.0 / numeric
          } else if (builtinId == BuiltinId.Gcd) {
            const integer = gcdLcmIntegerOrNaN(numeric)
            if (isNaN(integer)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            gcdValue = count == 0 ? integer : gcdPairCalc(gcdValue, integer)
          } else if (builtinId == BuiltinId.Lcm) {
            const integer = gcdLcmIntegerOrNaN(numeric)
            if (isNaN(integer)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
            lcmValue = count == 0 ? integer : lcmPairOrNaN(lcmValue, integer)
            if (isNaN(lcmValue)) {
              return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
            }
          } else if (builtinId == BuiltinId.Product) {
            product *= numeric
          } else if (builtinId == BuiltinId.Sumsq) {
            sumSquares += numeric * numeric
          }
          count += 1
        }
        continue
      }

      const numeric = directScalarNumberLikeText(
        slot,
        valueStack,
        tagStack,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (isNaN(numeric)) {
        if (
          tagStack[slot] == ValueTag.String ||
          builtinId == BuiltinId.Gcd ||
          builtinId == BuiltinId.Lcm ||
          builtinId == BuiltinId.Product ||
          builtinId == BuiltinId.Sumsq
        ) {
          return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        continue
      }
      if (builtinId == BuiltinId.Geomean) {
        if (numeric < 0.0) {
          return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (numeric == 0.0) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        logSum += Math.log(numeric)
      } else if (builtinId == BuiltinId.Harmean) {
        if (numeric <= 0.0) {
          return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        reciprocalSum += 1.0 / numeric
      } else if (builtinId == BuiltinId.Gcd) {
        const integer = gcdLcmIntegerOrNaN(numeric)
        if (isNaN(integer)) {
          return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        gcdValue = count == 0 ? integer : gcdPairCalc(gcdValue, integer)
      } else if (builtinId == BuiltinId.Lcm) {
        const integer = gcdLcmIntegerOrNaN(numeric)
        if (isNaN(integer)) {
          return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        lcmValue = count == 0 ? integer : lcmPairOrNaN(lcmValue, integer)
        if (isNaN(lcmValue)) {
          return writeAggregateError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      } else if (builtinId == BuiltinId.Product) {
        product *= numeric
      } else if (builtinId == BuiltinId.Sumsq) {
        sumSquares += numeric * numeric
      }
      count += 1
    }

    if (
      (builtinId == BuiltinId.Gcd || builtinId == BuiltinId.Lcm || builtinId == BuiltinId.Geomean || builtinId == BuiltinId.Harmean) &&
      count == 0
    ) {
      return writeAggregateError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let result = 0.0
    if (builtinId == BuiltinId.Gcd) {
      result = gcdValue
    } else if (builtinId == BuiltinId.Lcm) {
      result = lcmValue
    } else if (builtinId == BuiltinId.Product) {
      result = count == 0 ? 0.0 : product
    } else if (builtinId == BuiltinId.Geomean) {
      result = Math.exp(logSum / <f64>count)
    } else if (builtinId == BuiltinId.Harmean) {
      result = <f64>count / reciprocalSum
    } else {
      result = sumSquares
    }

    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  const criteriaConditionalResult = tryApplyCriteriaConditionalBuiltin(
    builtinId,
    argc,
    base,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
    stringOffsets,
    stringLengths,
    stringData,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )
  if (criteriaConditionalResult >= 0) {
    return criteriaConditionalResult
  }

  return -1
}
