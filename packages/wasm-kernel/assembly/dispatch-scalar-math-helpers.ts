import { ErrorCode, ValueTag } from './protocol'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
import { coerceScalarNumberLikeText } from './text-special'

export function writeScalarMathError(
  base: i32,
  error: ErrorCode,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

export function writeScalarMathNumber(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
}

export function writeScalarMathFiniteNumberOrNum(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return isFinite(value)
    ? writeScalarMathNumber(base, value, rangeIndexStack, valueStack, tagStack, kindStack)
    : writeScalarMathError(base, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
}

export function scalarMathNumberLikeText(
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

export function firstScalarMathError(base: i32, argc: i32, valueStack: Float64Array, tagStack: Uint8Array): ErrorCode {
  for (let index = 0; index < argc; index += 1) {
    const slot = base + index
    if (tagStack[slot] == ValueTag.Error) {
      return <ErrorCode>valueStack[slot]
    }
  }
  return ErrorCode.None
}
