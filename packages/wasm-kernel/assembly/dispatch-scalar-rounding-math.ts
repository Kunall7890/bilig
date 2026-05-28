import { BuiltinId, ErrorCode } from './protocol'
import {
  ceilingToMultiple,
  floorToMultiple,
  moduloValue,
  roundAwayFromZeroDigits,
  roundToDigits,
  roundTowardZeroDigits,
} from './numeric-core'
import { scalarMathNumberLikeText, writeScalarMathError, writeScalarMathNumber } from './dispatch-scalar-math-helpers'

export function tryApplyScalarRoundingMathBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Abs && argc == 1) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, Math.abs(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, roundToDigits(numeric, 0), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Round && argc == 2) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const digits = scalarMathNumberLikeText(
      base + 1,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric) || isNaN(digits)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, roundToDigits(numeric, <i32>digits), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, Math.floor(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Floor && argc == 2) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significance = scalarMathNumberLikeText(
      base + 1,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (isNaN(numeric) || isNaN(significance)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significance == 0.0) {
      return numeric == 0.0
        ? writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
        : writeScalarMathError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric > 0.0 && significance < 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, floorToMultiple(numeric, significance), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, Math.ceil(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ceiling && argc == 2) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significance = scalarMathNumberLikeText(
      base + 1,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (isNaN(numeric) || isNaN(significance)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significance == 0.0) {
      return numeric == 0.0
        ? writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
        : writeScalarMathError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (numeric > 0.0 && significance < 0.0) {
      return writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, ceilingToMultiple(numeric, significance), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.FloorMath && (argc == 1 || argc == 2 || argc == 3)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significanceRaw =
      argc >= 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 1.0
    const mode =
      argc == 3
        ? scalarMathNumberLikeText(
            base + 2,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(numeric) || isNaN(significanceRaw) || isNaN(mode)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significanceRaw == 0.0) {
      return writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const significance = Math.abs(significanceRaw)
    const result =
      numeric >= 0.0
        ? floorToMultiple(numeric, significance)
        : -(mode == 0.0 ? ceilingToMultiple(Math.abs(numeric), significance) : floorToMultiple(Math.abs(numeric), significance))
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.FloorPrecise && (argc == 1 || argc == 2)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significanceRaw =
      argc == 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 1.0
    if (isNaN(numeric) || isNaN(significanceRaw)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significanceRaw == 0.0) {
      return writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const significance = Math.abs(significanceRaw)
    return writeScalarMathNumber(base, floorToMultiple(numeric, significance), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.CeilingMath && (argc == 1 || argc == 2 || argc == 3)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significanceRaw =
      argc >= 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 1.0
    const mode =
      argc == 3
        ? scalarMathNumberLikeText(
            base + 2,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(numeric) || isNaN(significanceRaw) || isNaN(mode)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significanceRaw == 0.0) {
      return writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const significance = Math.abs(significanceRaw)
    const result =
      numeric >= 0.0
        ? ceilingToMultiple(numeric, significance)
        : -(mode == 0.0 ? floorToMultiple(Math.abs(numeric), significance) : ceilingToMultiple(Math.abs(numeric), significance))
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.CeilingPrecise || builtinId == BuiltinId.IsoCeiling) && (argc == 1 || argc == 2)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const significanceRaw =
      argc == 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 1.0
    if (isNaN(numeric) || isNaN(significanceRaw)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (significanceRaw == 0.0) {
      return writeScalarMathNumber(base, 0.0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const significance = Math.abs(significanceRaw)
    return writeScalarMathNumber(base, ceilingToMultiple(numeric, significance), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = scalarMathNumberLikeText(
      base + 1,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const dividend = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (isNaN(divisor) || isNaN(dividend)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (divisor == 0.0) {
      return writeScalarMathError(base, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeScalarMathNumber(base, moduloValue(dividend, divisor), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Int && argc == 1) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    return isNaN(numeric)
      ? writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeScalarMathNumber(base, Math.floor(numeric), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.RoundUp || builtinId == BuiltinId.RoundDown) && (argc == 1 || argc == 2)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const digitNumeric =
      argc == 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(numeric) || isNaN(digitNumeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const digits = <i32>digitNumeric
    const result = builtinId == BuiltinId.RoundUp ? roundAwayFromZeroDigits(numeric, digits) : roundTowardZeroDigits(numeric, digits)
    return writeScalarMathNumber(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Trunc && (argc == 1 || argc == 2)) {
    const numeric = scalarMathNumberLikeText(
      base,
      valueStack,
      tagStack,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const digitNumeric =
      argc == 2
        ? scalarMathNumberLikeText(
            base + 1,
            valueStack,
            tagStack,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0
    if (isNaN(numeric) || isNaN(digitNumeric)) {
      return writeScalarMathError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const digits = <i32>digitNumeric
    return writeScalarMathNumber(base, roundTowardZeroDigits(numeric, digits), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
