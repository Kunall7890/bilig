import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { coerceLogical } from './builtin-args'
import { compareSwitchValues } from './comparison'
import { BITWISE_NUM_ERROR, BITWISE_VALUE_ERROR, coerceBitwiseInteger, coerceBitwiseShift } from './numeric-core'
import { copySlotResult, STACK_KIND_RANGE, STACK_KIND_SCALAR, writeResult } from './result-io'

const IGNORED_REFERENCE_LOGICAL: i32 = i32.MIN_VALUE

function writeLogicInfoError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

function writeLogicInfoBoolean(
  base: i32,
  value: bool,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Boolean, value ? 1 : 0, rangeIndexStack, valueStack, tagStack, kindStack)
}

function writeLogicInfoNumber(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
}

function bitwiseErrorCode(value: i64): i32 {
  if (value == BITWISE_VALUE_ERROR) return ErrorCode.Value
  if (value == BITWISE_NUM_ERROR) return ErrorCode.Num
  return 0
}

function coerceReferenceLogical(tag: u8, value: f64, errorCode: i32): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Error) {
    return -errorCode - 1
  }
  return IGNORED_REFERENCE_LOGICAL
}

export function tryApplyLogicInfoBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellErrors: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if ((builtinId == BuiltinId.Bitand || builtinId == BuiltinId.Bitor || builtinId == BuiltinId.Bitxor) && argc == 2) {
    let accumulatorValue = coerceBitwiseInteger(tagStack[base], valueStack[base])
    let errorCode = bitwiseErrorCode(accumulatorValue)
    if (errorCode != 0) {
      return writeLogicInfoError(base, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let accumulator = accumulatorValue
    for (let index = 1; index < argc; index += 1) {
      const currentValue = coerceBitwiseInteger(tagStack[base + index], valueStack[base + index])
      errorCode = bitwiseErrorCode(currentValue)
      if (errorCode != 0) {
        return writeLogicInfoError(base, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (builtinId == BuiltinId.Bitand) {
        accumulator &= currentValue
      } else if (builtinId == BuiltinId.Bitor) {
        accumulator |= currentValue
      } else {
        accumulator ^= currentValue
      }
    }
    return writeLogicInfoNumber(base, <f64>accumulator, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Bitlshift || builtinId == BuiltinId.Bitrshift) && argc == 2) {
    const value = coerceBitwiseInteger(tagStack[base], valueStack[base])
    let errorCode = bitwiseErrorCode(value)
    if (errorCode != 0) {
      return writeLogicInfoError(base, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const shift = coerceBitwiseShift(tagStack[base + 1], valueStack[base + 1])
    errorCode = bitwiseErrorCode(shift)
    if (errorCode != 0) {
      return writeLogicInfoError(base, errorCode, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = <f64>value
    const shiftMagnitude = <f64>Math.abs(<f64>shift)
    const factor = Math.pow(2.0, shiftMagnitude)
    const shiftsLeft = builtinId == BuiltinId.Bitlshift ? shift >= 0 : shift < 0
    const result = shiftsLeft ? numeric * factor : Math.floor(numeric / factor)
    return writeLogicInfoNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.And) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let seenLogical = false
    let firstError = 0
    let result = true
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const coerced = coerceReferenceLogical(cellTags[memberIndex], cellNumbers[memberIndex], cellErrors[memberIndex])
          if (coerced == IGNORED_REFERENCE_LOGICAL) {
            continue
          }
          if (coerced < 0) {
            if (firstError == 0) {
              firstError = -coerced - 1
            }
            continue
          }
          seenLogical = true
          if (coerced == 0) {
            result = false
          }
        }
        continue
      }
      const coerced = tagStack[slot] == ValueTag.Empty ? IGNORED_REFERENCE_LOGICAL : coerceLogical(tagStack[slot], valueStack[slot])
      if (coerced == IGNORED_REFERENCE_LOGICAL) {
        continue
      }
      if (coerced < 0) {
        if (firstError == 0) {
          firstError = -coerced - 1
        }
        continue
      }
      seenLogical = true
      if (coerced == 0) {
        result = false
      }
    }
    if (firstError != 0) {
      return writeLogicInfoError(base, firstError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!seenLogical) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLogicInfoBoolean(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Or) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let seenLogical = false
    let firstError = 0
    let result = false
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const coerced = coerceReferenceLogical(cellTags[memberIndex], cellNumbers[memberIndex], cellErrors[memberIndex])
          if (coerced == IGNORED_REFERENCE_LOGICAL) {
            continue
          }
          if (coerced < 0) {
            if (firstError == 0) {
              firstError = -coerced - 1
            }
            continue
          }
          seenLogical = true
          if (coerced != 0) {
            result = true
          }
        }
        continue
      }
      const coerced = tagStack[slot] == ValueTag.Empty ? IGNORED_REFERENCE_LOGICAL : coerceLogical(tagStack[slot], valueStack[slot])
      if (coerced == IGNORED_REFERENCE_LOGICAL) {
        continue
      }
      if (coerced < 0) {
        if (firstError == 0) {
          firstError = -coerced - 1
        }
        continue
      }
      seenLogical = true
      if (coerced != 0) {
        result = true
      }
    }
    if (firstError != 0) {
      return writeLogicInfoError(base, firstError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!seenLogical) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLogicInfoBoolean(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Xor) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let seenLogical = false
    let parity = 0
    for (let index = 0; index < argc; index += 1) {
      const slot = base + index
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot]
        const start = rangeOffsets[rangeIndex]
        const length = <i32>rangeLengths[rangeIndex]
        for (let cursor = 0; cursor < length; cursor += 1) {
          const memberIndex = rangeMembers[start + cursor]
          const coerced = coerceReferenceLogical(cellTags[memberIndex], cellNumbers[memberIndex], cellErrors[memberIndex])
          if (coerced == IGNORED_REFERENCE_LOGICAL) {
            continue
          }
          if (coerced < 0) {
            return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
          }
          seenLogical = true
          parity = parity ^ (coerced != 0 ? 1 : 0)
        }
        continue
      }
      const coerced = tagStack[slot] == ValueTag.Empty ? IGNORED_REFERENCE_LOGICAL : coerceLogical(tagStack[slot], valueStack[slot])
      if (coerced == IGNORED_REFERENCE_LOGICAL) {
        continue
      }
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      seenLogical = true
      parity = parity ^ (coerced != 0 ? 1 : 0)
    }
    if (!seenLogical) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLogicInfoBoolean(base, parity != 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    const coerced = coerceLogical(tagStack[base], valueStack[base])
    if (coerced < 0) {
      return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLogicInfoBoolean(base, coerced == 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ifs) {
    if (argc < 2 || argc % 2 != 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 0; index < argc; index += 2) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index])
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (coerced != 0) {
        return copySlotResult(base, base + index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return writeLogicInfoError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Switch) {
    if (argc < 3) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeLogicInfoError(base, <i32>valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const hasDefault = (argc - 1) % 2 == 1
    const pairLimit = hasDefault ? argc - 1 : argc
    for (let index = 1; index < pairLimit; index += 2) {
      if (tagStack[base + index] == ValueTag.Error) {
        return writeLogicInfoError(base, <i32>valueStack[base + index], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const comparison = compareSwitchValues(
        tagStack[base],
        valueStack[base],
        tagStack[base + index],
        valueStack[base + index],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (comparison == 0) {
        return copySlotResult(base, base + index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return hasDefault
      ? copySlotResult(base, base + argc - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeLogicInfoError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsBlank && argc == 0) {
    return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsBlank && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.Empty, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsNumber && argc == 0) {
    return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsNumber && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.Number, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsText && argc == 0) {
    return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsText && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.String, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
