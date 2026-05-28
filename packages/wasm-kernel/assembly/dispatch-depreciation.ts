import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { dbDepreciation, ddbDepreciation, vdbDepreciation } from './date-finance'
import { toNumberExact } from './operands'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

function isNumberLikeTag(tag: u8): bool {
  return tag == ValueTag.Number || tag == ValueTag.Boolean || tag == ValueTag.Empty
}

function allNumberLikeArgs(base: i32, argc: i32, tagStack: Uint8Array): bool {
  for (let offset = 0; offset < argc; offset += 1) {
    if (!isNumberLikeTag(tagStack[base + offset])) {
      return false
    }
  }
  return true
}

function isDdbNumericDomainError(cost: f64, salvage: f64, life: f64, period: f64, factor: f64): bool {
  return (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(period) ||
    !isFinite(factor) ||
    cost <= 0.0 ||
    salvage < 0.0 ||
    life <= 0.0 ||
    period <= 0.0 ||
    period > life ||
    factor <= 0.0
  )
}

function isVdbNumericDomainError(cost: f64, salvage: f64, life: f64, startPeriod: f64, endPeriod: f64, factor: f64): bool {
  return (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(startPeriod) ||
    !isFinite(endPeriod) ||
    !isFinite(factor) ||
    cost <= 0.0 ||
    salvage < 0.0 ||
    life <= 0.0 ||
    startPeriod < 0.0 ||
    endPeriod < startPeriod ||
    factor <= 0.0
  )
}

export function tryApplyDepreciationBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (builtinId == BuiltinId.Db && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const month = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 12.0
    const maxPeriod = life + (month < 12.0 ? 1.0 : 0.0)
    if (
      !isFinite(life) ||
      !isFinite(period) ||
      !isFinite(month) ||
      life <= 0.0 ||
      period < 1.0 ||
      month < 1.0 ||
      month > 12.0 ||
      period > maxPeriod
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const depreciation = dbDepreciation(cost, salvage, life, period, month)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ddb && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!allNumberLikeArgs(base, argc, tagStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const factor = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 2.0
    if (isDdbNumericDomainError(cost, salvage, life, period, factor)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const depreciation = ddbDepreciation(cost, salvage, life, period, factor)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Vdb && argc >= 5 && argc <= 7) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numericArgc = argc >= 6 ? 6 : 5
    if (!allNumberLikeArgs(base, numericArgc, tagStack)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const startPeriod = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const endPeriod = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const factor = argc >= 6 ? toNumberExact(tagStack[base + 5], valueStack[base + 5]) : 2.0
    const noSwitch = argc >= 7 ? coerceBoolean(tagStack[base + 6], valueStack[base + 6]) : 0
    if (noSwitch < 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (isVdbNumericDomainError(cost, salvage, life, startPeriod, endPeriod, factor)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const depreciation = vdbDepreciation(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch != 0)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sln && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    if (isNaN(cost) || isNaN(salvage) || isNaN(life)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (life == 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (life < 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (cost - salvage) / life,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Syd && argc == 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(period)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (life <= 0.0 || period <= 0.0 || period > life) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const denominator = (life * (life + 1.0)) / 2.0
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      ((cost - salvage) * (life - period + 1.0)) / denominator,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
