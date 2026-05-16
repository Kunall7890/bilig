import { ValueTag, type CellValue } from '@bilig/protocol'
import { coerceNumber, firstError, integerValue, isErrorValue, numberResult, truncArg, valueError } from './cell-value-utils.js'
import { excelSerialWeekdayIndex, type ExcelDateSystem } from './excel-date.js'

type Builtin = (...args: CellValue[]) => CellValue

function isWeekendSerial(serial: number, dateSystem: ExcelDateSystem): boolean {
  const dow = excelSerialWeekdayIndex(serial, dateSystem)
  return dow === 0 || dow === 6
}

function weekendSerialDay(serial: number, dateSystem: ExcelDateSystem): number | undefined {
  return excelSerialWeekdayIndex(serial, dateSystem)
}

function weekendMaskFromCode(code: number): Set<number> | undefined {
  const twoDayWeekendMap: Record<number, readonly number[]> = {
    1: [6, 0],
    2: [0, 1],
    3: [1, 2],
    4: [2, 3],
    5: [3, 4],
    6: [4, 5],
    7: [5, 6],
  }
  if (code >= 1 && code <= 7) {
    return new Set(twoDayWeekendMap[code])
  }
  if (code >= 11 && code <= 17) {
    return new Set([(code - 10) % 7])
  }
  return undefined
}

function weekendMaskFromString(maskText: string): Set<number> | undefined {
  const trimmed = maskText.trim()
  if (!/^[01]{7}$/.test(trimmed) || trimmed === '1111111') {
    return undefined
  }
  const days = new Set<number>()
  for (let index = 0; index < trimmed.length; index += 1) {
    if (trimmed[index] !== '1') {
      continue
    }
    const dow = index === 6 ? 0 : index + 1
    days.add(dow)
  }
  return days
}

function normalizeWeekendMask(weekendArg: CellValue | undefined): Set<number> | CellValue {
  if (weekendArg === undefined) {
    return new Set([0, 6])
  }
  if (weekendArg.tag === ValueTag.Error) {
    return weekendArg
  }
  if (weekendArg.tag === ValueTag.String) {
    const mask = weekendMaskFromString(weekendArg.value)
    return mask ?? valueError()
  }
  const code = integerValue(weekendArg)
  if (code === undefined) {
    return valueError()
  }
  const mask = weekendMaskFromCode(code)
  return mask ?? valueError()
}

function normalizeHolidayDateSet(holidays: readonly CellValue[] | undefined): Set<number> | CellValue {
  if (!holidays || holidays.length === 0) {
    return new Set<number>()
  }

  const set = new Set<number>()
  for (const holiday of holidays) {
    const raw = coerceNumber(holiday)
    if (raw === undefined) {
      return valueError()
    }
    set.add(Math.trunc(raw))
  }
  return set
}

function isWeekendWithMask(serial: number, weekendDays: ReadonlySet<number>, dateSystem: ExcelDateSystem): boolean {
  const day = weekendSerialDay(serial, dateSystem)
  return day === undefined || weekendDays.has(day)
}

function offsetWorkday(start: number, offset: number, isWorkday: (serial: number) => boolean): number {
  let cursor = Math.trunc(start)
  const direction = offset >= 0 ? 1 : -1
  let remaining = Math.abs(offset)
  while (remaining > 0) {
    cursor += direction
    if (isWorkday(cursor)) {
      remaining -= 1
    }
  }
  return cursor
}

export function createWorkdayBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const offset = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof offset !== 'number') {
      return offset
    }

    const holidays = normalizeHolidayDateSet(args.slice(2))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value, dateSystem) && !holidays.has(Math.trunc(value))
    return numberResult(offsetWorkday(start, offset, isWorkday))
  }
}

export function createNetworkdaysBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const end = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof end !== 'number') {
      return end
    }

    const holidays = normalizeHolidayDateSet(args.slice(2))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value, dateSystem) && !holidays.has(Math.trunc(value))
    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = Math.trunc(start); ; cursor += step) {
      if (isWorkday(cursor)) {
        count += step
      }
      if (cursor === Math.trunc(end)) {
        break
      }
    }
    return numberResult(count)
  }
}

export function createWorkdayIntlBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const offset = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof offset !== 'number') {
      return offset
    }

    const weekendDays = normalizeWeekendMask(args[2])
    if (isErrorValue(weekendDays)) {
      return weekendDays
    }

    const holidays = normalizeHolidayDateSet(args.length <= 3 ? undefined : args.slice(3))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendWithMask(value, weekendDays, dateSystem) && !holidays.has(Math.trunc(value))
    return numberResult(offsetWorkday(start, offset, isWorkday))
  }
}

export function createNetworkdaysIntlBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const end = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof end !== 'number') {
      return end
    }

    const weekendDays = normalizeWeekendMask(args[2])
    if (isErrorValue(weekendDays)) {
      return weekendDays
    }

    const holidays = normalizeHolidayDateSet(args.length <= 3 ? undefined : args.slice(3))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendWithMask(value, weekendDays, dateSystem) && !holidays.has(Math.trunc(value))
    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = Math.trunc(start); ; cursor += step) {
      if (isWorkday(cursor)) {
        count += step
      }
      if (cursor === Math.trunc(end)) {
        break
      }
    }
    return numberResult(count)
  }
}
