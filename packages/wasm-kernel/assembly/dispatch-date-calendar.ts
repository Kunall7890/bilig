import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { scalarErrorAt } from './builtin-args'
import { coerceWeekendMask, isWorkdaySerial, isWorkdaySerialWithWeekendMask, validateHolidayArgument } from './calendar-workdays'
import {
  excelDateTextSerial,
  excelDays360Value,
  excelSerialWhole,
  excelWeeknumFromSerial,
  excelYearfracValue,
  isExcelDateSerialInRange,
  isExcelWeeknumReturnType,
} from './date-finance'
import { truncToInt } from './numeric-core'
import { scalarText } from './text-codec'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

const DAYS_VALUE_ERROR: i32 = i32.MIN_VALUE
const DAYS_NUM_ERROR: i32 = i32.MIN_VALUE + 1

export function tryApplyDateCalendarBuiltin(
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
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Days && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const end = daysDateSerial(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const start = daysDateSerial(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (end == DAYS_VALUE_ERROR || start == DAYS_VALUE_ERROR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (end == DAYS_NUM_ERROR || start == DAYS_NUM_ERROR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>(end - start), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Days360 && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base])
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const method = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0
    const value =
      startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE || (method != 0 && method != 1)
        ? NaN
        : excelDays360Value(startWhole, endWhole, method)
    if (isNaN(value)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Yearfrac && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const startWhole = excelSerialWhole(tagStack[base], valueStack[base])
    const endWhole = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const basis = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0
    if (basis < 0 || basis > 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        basis == i32.MIN_VALUE ? ErrorCode.Value : ErrorCode.Num,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const value = startWhole == i32.MIN_VALUE || endWhole == i32.MIN_VALUE ? NaN : excelYearfracValue(startWhole, endWhole, basis)
    if (isNaN(value)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Weeknum && (argc == 1 || argc == 2)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1
    if (returnType == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!isExcelWeeknumReturnType(returnType)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const serialWhole = excelSerialWhole(tagStack[base], valueStack[base])
    if (serialWhole == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (!isExcelDateSerialInRange(serialWhole)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Num, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const weeknum = excelWeeknumFromSerial(tagStack[base], valueStack[base], returnType)
    if (weeknum == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>weeknum, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Workday && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base])
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1])
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0
    if (
      !validateHolidayArgument(
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let cursor = start
    const direction = offset >= 0 ? 1 : -1
    let remaining = offset >= 0 ? offset : -offset
    while (remaining > 0) {
      cursor += direction
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (workday < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (workday == 1) {
        remaining -= 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>cursor, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Networkdays && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base])
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0

    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (workday < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (workday == 1) {
        count += step
      }
      if (cursor == end) {
        break
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.WorkdayIntl && argc >= 2 && argc <= 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base])
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1])
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0
    if (
      !validateHolidayArgument(
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let cursor = start
    const direction = offset >= 0 ? 1 : -1
    let remaining = offset >= 0 ? offset : -offset
    while (remaining > 0) {
      cursor += direction
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (workday < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (workday == 1) {
        remaining -= 1
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>cursor, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.NetworkdaysIntl && argc >= 2 && argc <= 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base])
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const weekendMask = coerceWeekendMask(
      argc >= 3,
      argc >= 3 ? tagStack[base + 2] : <u8>ValueTag.Empty,
      argc >= 3 ? valueStack[base + 2] : 0.0,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE || weekendMask == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const holidayKind = argc == 4 ? kindStack[base + 3] : STACK_KIND_SCALAR
    const holidayTag = argc == 4 ? tagStack[base + 3] : <u8>ValueTag.Empty
    const holidayValue = argc == 4 ? valueStack[base + 3] : 0.0
    const holidayRangeIndex = argc == 4 ? rangeIndexStack[base + 3] : 0

    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerialWithWeekendMask(
        cursor,
        weekendMask,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (workday < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (workday == 1) {
        count += step
      }
      if (cursor == end) {
        break
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>count, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}

function daysDateSerial(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (text == null) {
      return DAYS_VALUE_ERROR
    }
    const serial = excelDateTextSerial(text)
    return serial == i32.MIN_VALUE ? DAYS_VALUE_ERROR : serial
  }

  const serial = excelSerialWhole(tag, value)
  if (serial == i32.MIN_VALUE) {
    return DAYS_VALUE_ERROR
  }
  return isExcelDateSerialInRange(serial) ? serial : DAYS_NUM_ERROR
}
