import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'
import { coerceNumber, coerceText, firstError, integerValue, numberResult, truncArg, valueError } from './cell-value-utils.js'
import {
  MS_PER_DAY,
  addMonthsToExcelDate,
  daysInExcelMonth,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  excelSerialWeekdayIndex,
  floorDateSerial,
  isValidYearfracBasis,
  utcDateToExcelSerial,
  yearFracByBasis,
  type ExcelDateSystem,
  type ExcelDateParts,
} from './excel-date.js'
import { createBlockedBuiltinMap, datetimePlaceholderBuiltinNames } from './placeholder.js'
import {
  createNetworkdaysBuiltin,
  createNetworkdaysIntlBuiltin,
  createWorkdayBuiltin,
  createWorkdayIntlBuiltin,
} from './workday-builtins.js'
export {
  addMonthsToExcelDate,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  excelSerialWeekdayIndex,
  utcDateToExcelSerial,
} from './excel-date.js'
export type { ExcelDateParts, ExcelDateSystem } from './excel-date.js'

export type Builtin = (...args: CellValue[]) => CellValue
export type DateTimeProvider = () => Date
export type RandomProvider = () => number

const SECONDS_PER_DAY = 86_400
const MAX_EXCEL_DATE_SERIAL_BY_SYSTEM: Record<ExcelDateSystem, number> = {
  '1900': 2_958_465,
  '1904': 2_957_003,
}

function parseDateValueFromText(raw: string, dateSystem: ExcelDateSystem): number | undefined {
  const tokens = collectDateTextTokens(raw)
  if (tokens.length < 3) {
    return undefined
  }
  const parsed = parseDateTextTokens(tokens)
  if (parsed === undefined) {
    return undefined
  }
  return strictDatePartsToSerial(parsed.year, parsed.month, parsed.day, dateSystem)
}

function collectDateTextTokens(raw: string): string[] {
  const tokens: string[] = []
  let token = ''
  const trimmed = raw.trim()

  for (const char of trimmed) {
    const code = char.charCodeAt(0)
    const isDigit = code >= 48 && code <= 57
    const isLetter = (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
    if ((char === 'T' || char === 't') && tokens.length === 2 && token !== '' && /^\d+$/.test(token)) {
      tokens.push(token)
      break
    }
    if (isDigit || isLetter) {
      token += char
      continue
    }
    if (token !== '') {
      tokens.push(token)
      if (tokens.length === 3) {
        break
      }
      token = ''
    }
  }

  if (token !== '' && tokens.length < 3) {
    tokens.push(token)
  }
  return tokens
}

function parseDateTextTokens(tokens: readonly string[]): ExcelDateParts | undefined {
  const [first, second, third] = tokens
  if (first === undefined || second === undefined || third === undefined) {
    return undefined
  }

  const firstNumber = parseUnsignedDateToken(first)
  const secondNumber = parseUnsignedDateToken(second)
  const thirdNumber = parseUnsignedDateToken(third)
  const firstMonth = parseMonthNameToken(first)
  const secondMonth = parseMonthNameToken(second)

  if (firstNumber !== undefined && first.length === 4 && secondNumber !== undefined && thirdNumber !== undefined) {
    return { year: firstNumber, month: secondNumber, day: thirdNumber }
  }
  if (firstNumber !== undefined && secondMonth !== undefined && thirdNumber !== undefined) {
    return { year: normalizeDateTextYear(thirdNumber, third.length), month: secondMonth, day: firstNumber }
  }
  if (firstMonth !== undefined && secondNumber !== undefined && thirdNumber !== undefined) {
    return { year: normalizeDateTextYear(thirdNumber, third.length), month: firstMonth, day: secondNumber }
  }
  if (firstNumber !== undefined && secondNumber !== undefined && thirdNumber !== undefined) {
    return { year: normalizeDateTextYear(thirdNumber, third.length), month: firstNumber, day: secondNumber }
  }
  return undefined
}

function parseUnsignedDateToken(token: string): number | undefined {
  if (!/^\d+$/.test(token)) {
    return undefined
  }
  const value = Number(token)
  return Number.isSafeInteger(value) ? value : undefined
}

function parseMonthNameToken(token: string): number | undefined {
  const key = token.slice(0, 3).toUpperCase()
  switch (key) {
    case 'JAN':
      return 1
    case 'FEB':
      return 2
    case 'MAR':
      return 3
    case 'APR':
      return 4
    case 'MAY':
      return 5
    case 'JUN':
      return 6
    case 'JUL':
      return 7
    case 'AUG':
      return 8
    case 'SEP':
      return 9
    case 'OCT':
      return 10
    case 'NOV':
      return 11
    case 'DEC':
      return 12
    default:
      return undefined
  }
}

function normalizeDateTextYear(year: number, tokenLength: number): number {
  if (tokenLength <= 2) {
    return year <= 29 ? 2000 + year : 1900 + year
  }
  return year
}

function strictDatePartsToSerial(year: number, month: number, day: number, dateSystem: ExcelDateSystem): number | undefined {
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > daysInExcelMonth(year, month)) {
    return undefined
  }
  const serial = excelDatePartsToSerial(year, month, day, dateSystem)
  const roundTrip = serial === undefined ? undefined : excelSerialToDateParts(serial, dateSystem)
  if (serial === undefined || roundTrip === undefined || roundTrip.year !== year || roundTrip.month !== month || roundTrip.day !== day) {
    return undefined
  }
  return serial
}

function isValidDaysDateSerial(serial: number, dateSystem: ExcelDateSystem): boolean {
  return Number.isFinite(serial) && serial >= 0 && serial <= MAX_EXCEL_DATE_SERIAL_BY_SYSTEM[dateSystem]
}

function coerceDaysDateSerial(value: CellValue, dateSystem: ExcelDateSystem): number | CellValue {
  if (value.tag === ValueTag.String) {
    const serial = parseDateValueFromText(value.value, dateSystem)
    return serial === undefined ? valueError() : serial
  }
  const serial = truncArg(value)
  if (typeof serial !== 'number') {
    return serial
  }
  return isValidDaysDateSerial(serial, dateSystem) ? serial : numError()
}

export function parseTimeValueText(raw: string): number | undefined {
  const trimmed = raw.trim()
  const amPmMatch = trimmed.match(/^(.+?)\s+([aApP][mM])$/)
  const hasMeridiem = amPmMatch !== null
  const coreText = (hasMeridiem ? (amPmMatch?.[1] ?? '') : trimmed).trim()
  const firstColonIndex = coreText.indexOf(':')
  if (firstColonIndex < 0) {
    return undefined
  }
  let hourEndIndex = firstColonIndex - 1
  while (hourEndIndex >= 0 && /\s/.test(coreText[hourEndIndex] ?? '')) {
    hourEndIndex -= 1
  }
  let timeStartIndex = 0
  for (let index = hourEndIndex; index >= 0; index -= 1) {
    const char = coreText[index] ?? ''
    if (/\s/.test(char) || char === 'T' || char === 't') {
      timeStartIndex = index + 1
      break
    }
  }
  const timeText = coreText.slice(timeStartIndex).trim()
  const timeParts = timeText.split(':')
  if (timeParts.length < 2 || timeParts.length > 3) {
    return undefined
  }

  const [hoursText, minutesText, secondsText = '0'] = timeParts
  const hours = Number(hoursText)
  const minutes = Number(minutesText)
  const seconds = Number(secondsText)
  if (
    Number.isNaN(hours) ||
    Number.isNaN(minutes) ||
    Number.isNaN(seconds) ||
    !Number.isFinite(hours) ||
    !Number.isFinite(minutes) ||
    !Number.isFinite(seconds)
  ) {
    return undefined
  }

  const truncHours = Math.trunc(hours)
  const truncMinutes = Math.trunc(minutes)
  const truncSeconds = Math.trunc(seconds)
  const hasPm = hasMeridiem && amPmMatch?.[2]?.toLowerCase() === 'pm'

  if (truncMinutes < 0 || truncMinutes > 59 || truncSeconds < 0 || truncSeconds > 59) {
    return undefined
  }

  let hourValue = truncHours
  if (hasMeridiem) {
    if (truncHours < 1 || truncHours > 12) {
      return undefined
    }
    if (truncHours === 12) {
      hourValue = hasPm ? 12 : 0
    } else if (hasPm) {
      hourValue = truncHours + 12
    }
  } else if (truncHours === 24 && truncMinutes === 0 && truncSeconds === 0) {
    hourValue = 0
  } else if (truncHours < 0 || truncHours > 23) {
    return undefined
  }

  return (hourValue * 3600 + truncMinutes * 60 + truncSeconds) / SECONDS_PER_DAY
}

function createDays360Builtin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2 || args.length > 3) {
      return valueError()
    }

    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const method = args[2] === undefined ? 0 : integerValue(args[2], 0)
    if (method === undefined || (method !== 0 && method !== 1)) {
      return valueError()
    }

    if (typeof startSerial !== 'number') {
      return startSerial
    }
    if (typeof endSerial !== 'number') {
      return endSerial
    }
    const startParts = excelSerialToDateParts(startSerial, dateSystem)
    const endParts = excelSerialToDateParts(endSerial, dateSystem)
    if (!startParts || !endParts) {
      return valueError()
    }

    let startDay = startParts.day
    let endDay = endParts.day

    if (method === 0) {
      const startIsFebruaryMonthEnd = startParts.month === 2 && startDay === daysInExcelMonth(startParts.year, startParts.month)
      const endIsFebruaryMonthEnd = endParts.month === 2 && endDay === daysInExcelMonth(endParts.year, endParts.month)
      const startWasDayThirtyOrThirtyOne = startDay >= 30
      if (startDay === 31 || startIsFebruaryMonthEnd) {
        startDay = 30
      }
      if ((endDay === 31 && startWasDayThirtyOrThirtyOne) || (startIsFebruaryMonthEnd && endIsFebruaryMonthEnd)) {
        endDay = 30
      }
    } else {
      if (startDay === 31) {
        startDay = 30
      }
      if (endDay === 31) {
        endDay = 30
      }
    }

    return numberResult((endParts.year - startParts.year) * 360 + (endParts.month - startParts.month) * 30 + (endDay - startDay))
  }
}

function isoWeeknumFromDateParts(parts: ExcelDateParts): number {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
  const dow = date.getUTCDay()
  const dayShift = dow === 0 ? 7 : dow
  const shifted = new Date(date.getTime())
  shifted.setUTCDate(date.getUTCDate() + 4 - dayShift)
  const yearStart = new Date(Date.UTC(shifted.getUTCFullYear(), 0, 1))
  const dayOfYear = Math.floor((shifted.getTime() - yearStart.getTime()) / MS_PER_DAY) + 1
  return Math.floor((dayOfYear - 1) / 7) + 1
}

function createIsoWeeknumBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 1) {
      return valueError()
    }

    const serial = truncArg(args[0]!)
    if (typeof serial !== 'number') {
      return serial
    }

    const parts = excelSerialToDateParts(serial, dateSystem)
    if (parts === undefined) {
      return valueError()
    }

    return numberResult(isoWeeknumFromDateParts(parts))
  }
}

function createTimeValueBuiltin(): Builtin {
  return (value) => {
    if (value === undefined) {
      return valueError()
    }
    const error = firstError([value])
    if (error) {
      return error
    }
    const text = coerceText(value)
    if (text === undefined) {
      return valueError()
    }

    const parsed = parseTimeValueText(text)
    return parsed === undefined ? valueError() : numberResult(parsed)
  }
}

function createYearfracBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2 || args.length > 3) {
      return valueError()
    }

    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const basis = args[2] === undefined ? 0 : integerValue(args[2])
    if (typeof startSerial !== 'number' || typeof endSerial !== 'number' || basis === undefined || !isValidYearfracBasis(basis)) {
      return valueError()
    }

    const fraction = yearFracByBasis(startSerial, endSerial, basis, dateSystem)
    return fraction === undefined ? valueError() : numberResult(fraction)
  }
}

function normalizeSecondOfDay(serial: number): number | undefined {
  if (!Number.isFinite(serial) || serial < 0) {
    return undefined
  }
  const fraction = serial - floorDateSerial(serial)
  const normalizedFraction = fraction < 0 ? fraction + 1 : fraction
  return Math.floor(normalizedFraction * SECONDS_PER_DAY + 1e-9) % SECONDS_PER_DAY
}

function datedifValue(startSerial: number, endSerial: number, unit: string, dateSystem: ExcelDateSystem): number | undefined {
  if (startSerial > endSerial) {
    return undefined
  }
  const start = excelSerialToDateParts(startSerial, dateSystem)
  const end = excelSerialToDateParts(endSerial, dateSystem)
  if (!start || !end) {
    return undefined
  }
  const totalDays = Math.trunc(endSerial) - Math.trunc(startSerial)
  const totalMonths = (end.year - start.year) * 12 + (end.month - start.month) - (end.day < start.day ? 1 : 0)
  const totalYears = end.year - start.year - (end.month < start.month || (end.month === start.month && end.day < start.day) ? 1 : 0)

  switch (unit) {
    case 'D':
      return totalDays
    case 'M':
      return totalMonths
    case 'Y':
      return totalYears
    case 'YM':
      return ((totalMonths % 12) + 12) % 12
    case 'YD': {
      let comparisonYear = end.year
      let comparison = excelDatePartsToSerial(comparisonYear, start.month, start.day, dateSystem)
      if (comparison === undefined || comparison > endSerial) {
        comparisonYear -= 1
        comparison = excelDatePartsToSerial(comparisonYear, start.month, start.day, dateSystem)
      }
      return comparison === undefined ? undefined : Math.trunc(endSerial) - Math.trunc(comparison)
    }
    case 'MD':
      if (end.day >= start.day) {
        return end.day - start.day
      }
      return daysInExcelMonth(end.year, end.month === 1 ? 12 : end.month - 1) - start.day + end.day
    default:
      return undefined
  }
}

function normalizeTimeSerial(hours: number, minutes: number, seconds: number): number | undefined {
  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || !Number.isFinite(seconds)) {
    return undefined
  }
  if (hours < 0 || minutes < 0 || seconds < 0) {
    return undefined
  }
  if (hours > 32_767 || minutes > 32_767 || seconds > 32_767) {
    return undefined
  }
  const totalSeconds = Math.trunc(hours) * 3600 + Math.trunc(minutes) * 60 + Math.trunc(seconds)
  return (((totalSeconds % SECONDS_PER_DAY) + SECONDS_PER_DAY) % SECONDS_PER_DAY) / SECONDS_PER_DAY
}

function numError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Num }
}

export function createDateBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }

    const year = truncArg(args[0]!)
    const month = truncArg(args[1]!)
    const day = truncArg(args[2]!)
    if (typeof year !== 'number') return year
    if (typeof month !== 'number') return month
    if (typeof day !== 'number') return day

    const serial = excelDatePartsToSerial(year, month, day, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createDateValueBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (dateText) => {
    if (dateText === undefined) {
      return valueError()
    }
    const error = firstError([dateText])
    if (error) {
      return error
    }

    if (dateText.tag !== ValueTag.String) {
      return valueError()
    }

    const serial = parseDateValueFromText(dateText.value, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

function createDatePartBuiltin(part: keyof ExcelDateParts, dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (value) => {
    const error = firstError([value])
    if (error) {
      return error
    }

    const serial = coerceNumber(value)
    if (serial === undefined) {
      return valueError()
    }

    const parts = excelSerialToDateParts(serial, dateSystem)
    return parts ? numberResult(parts[part]) : valueError()
  }
}

function createTimeBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }

    const hour = truncArg(args[0]!)
    const minute = truncArg(args[1]!)
    const second = truncArg(args[2]!)
    if (typeof hour !== 'number') return hour
    if (typeof minute !== 'number') return minute
    if (typeof second !== 'number') return second

    if (!Number.isFinite(hour) || !Number.isFinite(minute) || !Number.isFinite(second)) {
      return valueError()
    }
    if (hour < 0 || minute < 0 || second < 0 || hour > 32_767 || minute > 32_767 || second > 32_767) {
      return numError()
    }

    const serial = normalizeTimeSerial(hour, minute, second)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

function createTimePartBuiltin(part: 'hour' | 'minute' | 'second'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    const [value] = args
    if (value === undefined) {
      return valueError()
    }

    const serial = coerceNumber(value)
    if (serial === undefined) {
      return valueError()
    }
    const seconds = normalizeSecondOfDay(serial)
    if (seconds === undefined) {
      return valueError()
    }

    switch (part) {
      case 'hour':
        return numberResult(Math.floor(seconds / 3600))
      case 'minute':
        return numberResult(Math.floor((seconds % 3600) / 60))
      case 'second':
        return numberResult(seconds % 60)
    }
  }
}

function createWeekdayBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 1 || args.length > 2) {
      return valueError()
    }
    const serial = coerceNumber(args[0]!)
    if (serial === undefined || serial < 0) {
      return valueError()
    }

    const weekdayIndex = excelSerialWeekdayIndex(serial, dateSystem)
    if (weekdayIndex === undefined) {
      return valueError()
    }
    const sundayOne = weekdayIndex + 1
    if (args.length === 1) {
      return numberResult(sundayOne)
    }

    const returnType = truncArg(args[1]!)
    if (typeof returnType !== 'number') {
      return returnType
    }
    if (returnType === 3) {
      return numberResult(sundayOne === 1 ? 6 : sundayOne - 2)
    }

    const startDayMap: Record<number, number> = {
      1: 1,
      2: 2,
      11: 2,
      12: 3,
      13: 4,
      14: 5,
      15: 6,
      16: 7,
      17: 1,
    }
    const startDay = startDayMap[returnType]
    if (startDay === undefined) {
      return valueError()
    }
    return numberResult(((sundayOne - startDay + 7) % 7) + 1)
  }
}

function createDaysBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 2) {
      return valueError()
    }
    const endSerial = coerceDaysDateSerial(args[0]!, dateSystem)
    const startSerial = coerceDaysDateSerial(args[1]!, dateSystem)
    if (typeof endSerial !== 'number') {
      return endSerial
    }
    if (typeof startSerial !== 'number') {
      return startSerial
    }
    return numberResult(endSerial - startSerial)
  }
}

function createWeeknumBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 1 || args.length > 2) {
      return valueError()
    }

    const serial = truncArg(args[0]!)
    if (typeof serial !== 'number') {
      return serial
    }

    const returnType = args[1] === undefined ? 1 : truncArg(args[1])
    if (typeof returnType !== 'number') {
      return returnType
    }

    const dateParts = excelSerialToDateParts(serial, dateSystem)
    if (!dateParts) {
      return valueError()
    }
    if (returnType === 21) {
      return numberResult(isoWeeknumFromDateParts(dateParts))
    }

    let weekStartDay: number
    if (returnType === 1 || returnType === 17) {
      weekStartDay = 0
    } else if (returnType === 2 || returnType === 11) {
      weekStartDay = 1
    } else if (returnType === 12) {
      weekStartDay = 2
    } else if (returnType === 13) {
      weekStartDay = 3
    } else if (returnType === 14) {
      weekStartDay = 4
    } else if (returnType === 15) {
      weekStartDay = 5
    } else if (returnType === 16) {
      weekStartDay = 6
    } else {
      return valueError()
    }

    const serialJan1 = excelDatePartsToSerial(dateParts.year, 1, 1, dateSystem)
    if (serialJan1 === undefined) {
      return valueError()
    }

    const jan1Weekday = excelSerialWeekdayIndex(serialJan1, dateSystem)
    if (jan1Weekday === undefined) {
      return valueError()
    }
    const shift = (jan1Weekday - weekStartDay + 7) % 7

    let dayOfYear = dateParts.day
    for (let month = 1; month < dateParts.month; month += 1) {
      dayOfYear += daysInExcelMonth(dateParts.year, month)
    }

    return numberResult(Math.floor((dayOfYear - 1 + shift) / 7) + 1)
  }
}

export function createTodayBuiltin(now: DateTimeProvider = () => new Date(), dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }
    return numberResult(Math.floor(utcDateToExcelSerial(now(), dateSystem)))
  }
}

export function createNowBuiltin(now: DateTimeProvider = () => new Date(), dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }
    return numberResult(utcDateToExcelSerial(now(), dateSystem))
  }
}

export function createRandBuiltin(random: RandomProvider = () => Math.random()): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }

    const next = random()
    if (!Number.isFinite(next)) {
      return valueError()
    }

    const bounded = Math.min(Math.max(next, 0), 1 - Number.EPSILON)
    return numberResult(bounded)
  }
}

export function createEdateBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months])
    if (error) {
      return error
    }

    const startSerial = coerceNumber(startDate)
    const monthOffset = truncArg(months)
    if (startSerial === undefined) {
      return valueError()
    }
    if (typeof monthOffset !== 'number') {
      return monthOffset
    }

    const serial = addMonthsToExcelDate(startSerial, monthOffset, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createEomonthBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months])
    if (error) {
      return error
    }

    const startSerial = coerceNumber(startDate)
    const monthOffset = truncArg(months)
    if (startSerial === undefined) {
      return valueError()
    }
    if (typeof monthOffset !== 'number') {
      return monthOffset
    }

    const serial = endOfMonthExcelDate(startSerial, monthOffset, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createDatedifBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }
    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const unit = coerceText(args[2]!)?.trim().toUpperCase()
    if (typeof startSerial !== 'number' || typeof endSerial !== 'number' || !unit) {
      return valueError()
    }
    const value = datedifValue(startSerial, endSerial, unit, dateSystem)
    return value === undefined ? valueError() : numberResult(value)
  }
}

const datetimePlaceholderBuiltins = createBlockedBuiltinMap(datetimePlaceholderBuiltinNames)

export function createDateTimeBuiltins(dateSystem: ExcelDateSystem = '1900'): Record<string, Builtin> {
  return {
    DATE: createDateBuiltin(dateSystem),
    DATEVALUE: createDateValueBuiltin(dateSystem),
    YEAR: createDatePartBuiltin('year', dateSystem),
    MONTH: createDatePartBuiltin('month', dateSystem),
    DAY: createDatePartBuiltin('day', dateSystem),
    TIME: createTimeBuiltin(),
    HOUR: createTimePartBuiltin('hour'),
    MINUTE: createTimePartBuiltin('minute'),
    SECOND: createTimePartBuiltin('second'),
    WEEKDAY: createWeekdayBuiltin(dateSystem),
    DAYS: createDaysBuiltin(dateSystem),
    WEEKNUM: createWeeknumBuiltin(dateSystem),
    DAYS360: createDays360Builtin(dateSystem),
    ISOWEEKNUM: createIsoWeeknumBuiltin(dateSystem),
    TIMEVALUE: createTimeValueBuiltin(),
    YEARFRAC: createYearfracBuiltin(dateSystem),
    WORKDAY: createWorkdayBuiltin(dateSystem),
    'WORKDAY.INTL': createWorkdayIntlBuiltin(dateSystem),
    NETWORKDAYS: createNetworkdaysBuiltin(dateSystem),
    'NETWORKDAYS.INTL': createNetworkdaysIntlBuiltin(dateSystem),
    TODAY: createTodayBuiltin(() => new Date(), dateSystem),
    NOW: createNowBuiltin(() => new Date(), dateSystem),
    RAND: createRandBuiltin(),
    EDATE: createEdateBuiltin(dateSystem),
    EOMONTH: createEomonthBuiltin(dateSystem),
    DATEDIF: createDatedifBuiltin(dateSystem),
    ...datetimePlaceholderBuiltins,
  }
}

export const datetimeBuiltins: Record<string, Builtin> = createDateTimeBuiltins()
