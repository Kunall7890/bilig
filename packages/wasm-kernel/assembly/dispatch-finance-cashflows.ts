import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import {
  cumulativePeriodicPaymentCalc,
  futureValueCalc,
  periodicPaymentCalc,
  presentValueCalc,
  principalPaymentCalc,
  interestPaymentCalc,
  solveRateCalc,
  totalPeriodsCalc,
} from './cashflows'
import { collectNumericValuesFromArgsWithText, orderStatisticErrorCode } from './statistics-tests'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'
import { coerceScalarNumberLikeText } from './text-special'

class CashflowArgumentContext {
  valueStack: Float64Array
  tagStack: Uint8Array
  stringOffsets: Uint32Array
  stringLengths: Uint32Array
  stringData: Uint16Array
  outputStringOffsets: Uint32Array
  outputStringLengths: Uint32Array
  outputStringData: Uint16Array

  constructor(
    valueStack: Float64Array,
    tagStack: Uint8Array,
    stringOffsets: Uint32Array,
    stringLengths: Uint32Array,
    stringData: Uint16Array,
    outputStringOffsets: Uint32Array,
    outputStringLengths: Uint32Array,
    outputStringData: Uint16Array,
  ) {
    this.valueStack = valueStack
    this.tagStack = tagStack
    this.stringOffsets = stringOffsets
    this.stringLengths = stringLengths
    this.stringData = stringData
    this.outputStringOffsets = outputStringOffsets
    this.outputStringLengths = outputStringLengths
    this.outputStringData = outputStringData
  }

  numberAt(slot: i32): f64 {
    return coerceScalarNumberLikeText(
      this.tagStack[slot],
      this.valueStack[slot],
      this.stringOffsets,
      this.stringLengths,
      this.stringData,
      this.outputStringOffsets,
      this.outputStringLengths,
      this.outputStringData,
    )
  }

  paymentTypeAt(slot: i32): i32 {
    const numeric = this.numberAt(slot)
    if (!isFinite(numeric)) {
      return -1
    }
    const type = <i32>numeric
    return type == 0 || type == 1 ? type : -1
  }
}

export function tryApplyFinanceCashflowBuiltin(
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
  const args = new CashflowArgumentContext(
    valueStack,
    tagStack,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  )

  if (builtinId == BuiltinId.Pv && argc >= 3 && argc <= 5) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const periods = args.numberAt(base + 1)
    const payment = args.numberAt(base + 2)
    const future = argc >= 4 ? args.numberAt(base + 3) : 0.0
    const paymentTypeValue = argc >= 5 ? args.paymentTypeAt(base + 4) : 0
    const present = paymentTypeValue < 0 ? NaN : presentValueCalc(rate, periods, payment, future, paymentTypeValue)
    return isNaN(present)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, present, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Pmt && argc >= 3 && argc <= 5) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const periods = args.numberAt(base + 1)
    const present = args.numberAt(base + 2)
    const future = argc >= 4 ? args.numberAt(base + 3) : 0.0
    const paymentTypeValue = argc >= 5 ? args.paymentTypeAt(base + 4) : 0
    const payment = paymentTypeValue < 0 ? NaN : periodicPaymentCalc(rate, periods, present, future, paymentTypeValue)
    return isNaN(payment)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, payment, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Nper && argc >= 3 && argc <= 5) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const payment = args.numberAt(base + 1)
    const present = args.numberAt(base + 2)
    const future = argc >= 4 ? args.numberAt(base + 3) : 0.0
    const paymentTypeValue = argc >= 5 ? args.paymentTypeAt(base + 4) : 0
    const periods = paymentTypeValue < 0 ? NaN : totalPeriodsCalc(rate, payment, present, future, paymentTypeValue)
    return isNaN(periods)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, periods, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Npv && argc >= 2) {
    const rate = args.numberAt(base)
    if (isNaN(rate)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const values = collectNumericValuesFromArgsWithText(
      base + 1,
      argc - 1,
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
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (values.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let total = 0.0
    for (let index = 0; index < values.length; index += 1) {
      total += unchecked(values[index]) / Math.pow(1.0 + rate, <f64>(index + 1))
    }
    return !isFinite(total)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, total, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Rate && argc >= 3 && argc <= 6) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const periods = args.numberAt(base)
    const payment = args.numberAt(base + 1)
    const present = args.numberAt(base + 2)
    const future = argc >= 4 ? args.numberAt(base + 3) : 0.0
    const paymentTypeValue = argc >= 5 ? args.paymentTypeAt(base + 4) : 0
    const guess = argc >= 6 ? args.numberAt(base + 5) : 0.1
    const rate = paymentTypeValue < 0 ? NaN : solveRateCalc(periods, payment, present, future, paymentTypeValue, guess)
    return isNaN(rate)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, rate, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Ipmt || builtinId == BuiltinId.Ppmt) && argc >= 4 && argc <= 6) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const period = args.numberAt(base + 1)
    const periods = args.numberAt(base + 2)
    const present = args.numberAt(base + 3)
    const future = argc >= 5 ? args.numberAt(base + 4) : 0.0
    const paymentTypeValue = argc >= 6 ? args.paymentTypeAt(base + 5) : 0
    const result =
      paymentTypeValue < 0
        ? NaN
        : builtinId == BuiltinId.Ipmt
          ? interestPaymentCalc(rate, period, periods, present, future, paymentTypeValue)
          : principalPaymentCalc(rate, period, periods, present, future, paymentTypeValue)
    return isNaN(result)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ispmt && argc == 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const period = args.numberAt(base + 1)
    const periods = args.numberAt(base + 2)
    const present = args.numberAt(base + 3)
    if (isNaN(rate) || isNaN(period) || isNaN(periods) || isNaN(present) || periods <= 0.0 || period < 1.0 || period > periods) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      present * rate * (period / periods - 1.0),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.Cumipmt || builtinId == BuiltinId.Cumprinc) && argc == 6) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const periods = args.numberAt(base + 1)
    const present = args.numberAt(base + 2)
    const startPeriodNumeric = args.numberAt(base + 3)
    const endPeriodNumeric = args.numberAt(base + 4)
    const paymentTypeNumeric = args.numberAt(base + 5)
    if (
      isNaN(rate) ||
      isNaN(periods) ||
      isNaN(present) ||
      isNaN(startPeriodNumeric) ||
      isNaN(endPeriodNumeric) ||
      isNaN(paymentTypeNumeric)
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const startPeriod = <i32>startPeriodNumeric
    const endPeriod = <i32>endPeriodNumeric
    const paymentTypeValue = <i32>paymentTypeNumeric
    if (
      !isFinite(rate) ||
      !isFinite(periods) ||
      !isFinite(present) ||
      !isFinite(startPeriodNumeric) ||
      !isFinite(endPeriodNumeric) ||
      !isFinite(paymentTypeNumeric) ||
      rate <= 0.0 ||
      periods <= 0.0 ||
      present <= 0.0 ||
      startPeriod < 1 ||
      endPeriod < 1 ||
      startPeriod > endPeriod ||
      paymentTypeValue < 0 ||
      paymentTypeValue > 1
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const total = cumulativePeriodicPaymentCalc(
      rate,
      periods,
      present,
      startPeriod,
      endPeriod,
      paymentTypeValue,
      builtinId == BuiltinId.Cumprinc,
    )
    return isNaN(total)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, total, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Fv && argc >= 3 && argc <= 5) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const periods = args.numberAt(base + 1)
    const payment = args.numberAt(base + 2)
    const present = argc >= 4 ? args.numberAt(base + 3) : 0.0
    const paymentTypeValue = argc >= 5 ? args.paymentTypeAt(base + 4) : 0
    const future = paymentTypeValue < 0 ? NaN : futureValueCalc(rate, periods, payment, present, paymentTypeValue)
    return isNaN(future)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, future, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Fvschedule && argc >= 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const principal = args.numberAt(base)
    if (isNaN(principal)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let result = principal
    for (let index = 1; index < argc; index += 1) {
      const rate = args.numberAt(base + index)
      if (isNaN(rate)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result *= 1.0 + rate
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Effect || builtinId == BuiltinId.Nominal) && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const periodsNumeric = args.numberAt(base + 1)
    const periods = Math.trunc(periodsNumeric)
    if (isNaN(rate) || isNaN(periodsNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!isFinite(rate) || !isFinite(periodsNumeric) || periods < 1.0 || rate <= 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result =
      builtinId == BuiltinId.Effect ? Math.pow(1.0 + rate / periods, periods) - 1.0 : periods * (Math.pow(1.0 + rate, 1.0 / periods) - 1.0)
    return !isFinite(result)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Pduration && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rate = args.numberAt(base)
    const present = args.numberAt(base + 1)
    const future = args.numberAt(base + 2)
    if (isNaN(rate) || isNaN(present) || isNaN(future)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!isFinite(rate) || !isFinite(present) || !isFinite(future) || rate <= 0.0 || present <= 0.0 || future <= 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.log(future / present) / Math.log(1.0 + rate)
    return !isFinite(result)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Rri && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const periods = args.numberAt(base)
    const present = args.numberAt(base + 1)
    const future = args.numberAt(base + 2)
    if (isNaN(periods) || isNaN(present) || isNaN(future)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!isFinite(periods) || !isFinite(present) || !isFinite(future) || periods <= 0.0 || present == 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = Math.pow(future / present, 1.0 / periods) - 1.0
    return !isFinite(result)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
