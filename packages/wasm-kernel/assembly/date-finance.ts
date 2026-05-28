import { ValueTag } from './protocol'

function toNumberExact(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value
  if (tag == ValueTag.Empty) return 0
  return NaN
}

function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric
}

function floorDiv(a: i32, b: i32): i32 {
  let quotient = a / b
  const remainder = a % b
  if (remainder != 0 && remainder > 0 != b > 0) {
    quotient -= 1
  }
  return quotient
}

function daysFromCivil(year: i32, month: i32, day: i32): i32 {
  let adjustedYear = year
  if (month <= 2) {
    adjustedYear -= 1
  }
  const era = adjustedYear >= 0 ? adjustedYear / 400 : (adjustedYear - 399) / 400
  const yearOfEra = adjustedYear - era * 400
  const shiftedMonth = month + (month > 2 ? -3 : 9)
  const dayOfYear = (153 * shiftedMonth + 2) / 5 + day - 1
  const dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear
  return era * 146097 + dayOfEra - 719468
}

function civilYear(days: i32): i32 {
  const shifted = days + 719468
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097
  const dayOfEra = shifted - era * 146097
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100)
  const monthPrime = (5 * dayOfYear + 2) / 153
  const month = monthPrime + (monthPrime < 10 ? 3 : -9)
  return yearOfEra + era * 400 + (month <= 2 ? 1 : 0)
}

function civilMonth(days: i32): i32 {
  const shifted = days + 719468
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097
  const dayOfEra = shifted - era * 146097
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100)
  const monthPrime = (5 * dayOfYear + 2) / 153
  return monthPrime + (monthPrime < 10 ? 3 : -9)
}

function civilDay(days: i32): i32 {
  const shifted = days + 719468
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097
  const dayOfEra = shifted - era * 146097
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100)
  const monthPrime = (5 * dayOfYear + 2) / 153
  return dayOfYear - (153 * monthPrime + 2) / 5 + 1
}

const EXCEL_EPOCH_DAYS: i32 = -25568
const EXCEL_MAX_DATE_SERIAL_1900: i32 = 2958465
export const EXCEL_SECONDS_PER_DAY: i32 = 86400

export function excelSerialWhole(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value)
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>Math.floor(numeric)
}

export function isExcelDateSerialInRange(serial: i32): bool {
  return serial >= 0 && serial <= EXCEL_MAX_DATE_SERIAL_1900
}

export function isExcelWeekdayReturnType(returnType: i32): bool {
  return (
    returnType == 1 ||
    returnType == 2 ||
    returnType == 3 ||
    returnType == 11 ||
    returnType == 12 ||
    returnType == 13 ||
    returnType == 14 ||
    returnType == 15 ||
    returnType == 16 ||
    returnType == 17
  )
}

export function isExcelWeeknumReturnType(returnType: i32): bool {
  return (
    returnType == 1 ||
    returnType == 2 ||
    returnType == 11 ||
    returnType == 12 ||
    returnType == 13 ||
    returnType == 14 ||
    returnType == 15 ||
    returnType == 16 ||
    returnType == 17 ||
    returnType == 21
  )
}

function daysInExcelMonth(year: i32, month: i32): i32 {
  if (year == 1900 && month == 2) {
    return 29
  }
  const start = daysFromCivil(year, month, 1)
  const nextMonth = month == 12 ? 1 : month + 1
  const nextYear = month == 12 ? year + 1 : year
  const end = daysFromCivil(nextYear, nextMonth, 1)
  return end - start
}

function strictExcelDateSerial(year: i32, month: i32, day: i32): i32 {
  if (year < 1900 || year > 9999 || month < 1 || month > 12) {
    return i32.MIN_VALUE
  }
  if (day < 1 || day > daysInExcelMonth(year, month)) {
    return i32.MIN_VALUE
  }
  if (year == 1900 && month == 2 && day == 29) {
    return 60
  }

  const days = daysFromCivil(year, month, day)
  let serial = days - EXCEL_EPOCH_DAYS
  if (serial >= 60) {
    serial += 1
  }
  return isExcelDateSerialInRange(serial) ? serial : i32.MIN_VALUE
}

function trimAscii(input: string): string {
  let start = 0
  let end = input.length
  while (start < end && input.charCodeAt(start) <= 32) {
    start += 1
  }
  while (end > start && input.charCodeAt(end - 1) <= 32) {
    end -= 1
  }
  return input.slice(start, end)
}

function isDigitCode(code: i32): bool {
  return code >= 48 && code <= 57
}

function isLetterCode(code: i32): bool {
  return (code >= 65 && code <= 90) || (code >= 97 && code <= 122)
}

function tokenIsDigits(token: string): bool {
  if (token.length == 0) {
    return false
  }
  for (let index = 0; index < token.length; index += 1) {
    if (!isDigitCode(token.charCodeAt(index))) {
      return false
    }
  }
  return true
}

function collectDateTextTokens(raw: string): Array<string> {
  const tokens = new Array<string>()
  const trimmed = trimAscii(raw)
  let token = ''
  for (let index = 0; index < trimmed.length; index += 1) {
    const code = trimmed.charCodeAt(index)
    if ((code == 84 || code == 116) && tokens.length == 2 && token.length > 0 && tokenIsDigits(token)) {
      tokens.push(token)
      break
    }
    if (isDigitCode(code) || isLetterCode(code)) {
      token += String.fromCharCode(code)
      continue
    }
    if (token.length > 0) {
      tokens.push(token)
      if (tokens.length == 3) {
        break
      }
      token = ''
    }
  }
  if (token.length > 0 && tokens.length < 3) {
    tokens.push(token)
  }
  return tokens
}

function parseUnsignedDateToken(token: string): i32 {
  if (!tokenIsDigits(token)) {
    return i32.MIN_VALUE
  }
  let value: i32 = 0
  for (let index = 0; index < token.length; index += 1) {
    value = value * 10 + (token.charCodeAt(index) - 48)
    if (value > 9999) {
      return i32.MIN_VALUE
    }
  }
  return value
}

function upperAscii(code: i32): i32 {
  return code >= 97 && code <= 122 ? code - 32 : code
}

function monthFromToken(token: string): i32 {
  if (token.length < 3) {
    return i32.MIN_VALUE
  }
  const first = upperAscii(token.charCodeAt(0))
  const second = upperAscii(token.charCodeAt(1))
  const third = upperAscii(token.charCodeAt(2))
  if (first == 74 && second == 65 && third == 78) return 1
  if (first == 70 && second == 69 && third == 66) return 2
  if (first == 77 && second == 65 && third == 82) return 3
  if (first == 65 && second == 80 && third == 82) return 4
  if (first == 77 && second == 65 && third == 89) return 5
  if (first == 74 && second == 85 && third == 78) return 6
  if (first == 74 && second == 85 && third == 76) return 7
  if (first == 65 && second == 85 && third == 71) return 8
  if (first == 83 && second == 69 && third == 80) return 9
  if (first == 79 && second == 67 && third == 84) return 10
  if (first == 78 && second == 79 && third == 86) return 11
  if (first == 68 && second == 69 && third == 67) return 12
  return i32.MIN_VALUE
}

function normalizeDateTextYear(year: i32, tokenLength: i32): i32 {
  if (tokenLength <= 2) {
    return year <= 29 ? 2000 + year : 1900 + year
  }
  return year
}

export function excelDateTextSerial(raw: string): i32 {
  const tokens = collectDateTextTokens(raw)
  if (tokens.length < 3) {
    return i32.MIN_VALUE
  }

  const first = tokens[0]
  const second = tokens[1]
  const third = tokens[2]
  const firstNumber = parseUnsignedDateToken(first)
  const secondNumber = parseUnsignedDateToken(second)
  const thirdNumber = parseUnsignedDateToken(third)
  const firstMonth = monthFromToken(first)
  const secondMonth = monthFromToken(second)

  if (firstNumber != i32.MIN_VALUE && first.length == 4 && secondNumber != i32.MIN_VALUE && thirdNumber != i32.MIN_VALUE) {
    return strictExcelDateSerial(firstNumber, secondNumber, thirdNumber)
  }
  if (firstNumber != i32.MIN_VALUE && secondMonth != i32.MIN_VALUE && thirdNumber != i32.MIN_VALUE) {
    return strictExcelDateSerial(normalizeDateTextYear(thirdNumber, third.length), secondMonth, firstNumber)
  }
  if (firstMonth != i32.MIN_VALUE && secondNumber != i32.MIN_VALUE && thirdNumber != i32.MIN_VALUE) {
    return strictExcelDateSerial(normalizeDateTextYear(thirdNumber, third.length), firstMonth, secondNumber)
  }
  if (firstNumber != i32.MIN_VALUE && secondNumber != i32.MIN_VALUE && thirdNumber != i32.MIN_VALUE) {
    return strictExcelDateSerial(normalizeDateTextYear(thirdNumber, third.length), firstNumber, secondNumber)
  }
  return i32.MIN_VALUE
}

function isLeapYear(year: i32): bool {
  return year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

export function excelYearPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE
  }
  if (whole == 60) {
    return 1900
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1
  return civilYear(EXCEL_EPOCH_DAYS + adjustedWhole)
}

export function excelMonthPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE
  }
  if (whole == 60) {
    return 2
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1
  return civilMonth(EXCEL_EPOCH_DAYS + adjustedWhole)
}

export function excelDayPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE
  }
  if (whole == 60) {
    return 29
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1
  return civilDay(EXCEL_EPOCH_DAYS + adjustedWhole)
}

export function excelTimeSerialFromNumbers(hourNumeric: f64, minuteNumeric: f64, secondNumeric: f64): f64 {
  if (!isFinite(hourNumeric) || !isFinite(minuteNumeric) || !isFinite(secondNumeric)) {
    return NaN
  }
  const hour = <f64>(<i32>hourNumeric)
  const minute = <f64>(<i32>minuteNumeric)
  const second = <f64>(<i32>secondNumeric)
  if (hour < 0 || minute < 0 || second < 0 || hour > 32767 || minute > 32767 || second > 32767) {
    return NaN
  }
  let totalSeconds = hour * 3600.0 + minute * 60.0 + second
  totalSeconds %= <f64>EXCEL_SECONDS_PER_DAY
  if (totalSeconds < 0) {
    totalSeconds += <f64>EXCEL_SECONDS_PER_DAY
  }
  return totalSeconds / <f64>EXCEL_SECONDS_PER_DAY
}

export function excelTimeSerial(hourTag: u8, hourValue: f64, minuteTag: u8, minuteValue: f64, secondTag: u8, secondValue: f64): f64 {
  return excelTimeSerialFromNumbers(
    toNumberExact(hourTag, hourValue),
    toNumberExact(minuteTag, minuteValue),
    toNumberExact(secondTag, secondValue),
  )
}

export function excelFloorSecondOfDayFromNumber(numeric: f64): i32 {
  if (isNaN(numeric) || numeric < 0) {
    return i32.MIN_VALUE
  }
  const whole = Math.floor(numeric)
  let fraction = numeric - whole
  if (fraction < 0) {
    fraction += 1.0
  }
  let seconds = <i32>Math.floor(fraction * <f64>EXCEL_SECONDS_PER_DAY + 1e-9)
  if (seconds >= EXCEL_SECONDS_PER_DAY) {
    seconds = 0
  }
  return seconds
}

export function excelRoundedSecondOfDayFromNumber(numeric: f64): i32 {
  if (isNaN(numeric) || numeric < 0) {
    return i32.MIN_VALUE
  }
  const whole = Math.floor(numeric)
  let fraction = numeric - whole
  if (fraction < 0) {
    fraction += 1.0
  }
  let seconds = <i32>Math.floor(fraction * <f64>EXCEL_SECONDS_PER_DAY + 0.5)
  if (seconds >= EXCEL_SECONDS_PER_DAY) {
    seconds %= EXCEL_SECONDS_PER_DAY
  }
  return seconds
}

export function excelSecondOfDay(tag: u8, value: f64): i32 {
  return excelFloorSecondOfDayFromNumber(toNumberExact(tag, value))
}

export function excelRoundedSecondOfDay(tag: u8, value: f64): i32 {
  return excelRoundedSecondOfDayFromNumber(toNumberExact(tag, value))
}

export function excelWeekdayFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE || !isExcelDateSerialInRange(whole)) {
    return i32.MIN_VALUE
  }
  const adjustedWhole = whole - 1
  const sundayOne = (((adjustedWhole % 7) + 7) % 7) + 1
  if (returnType == 1) {
    return sundayOne
  }
  if (returnType == 3) {
    return sundayOne == 1 ? 6 : sundayOne - 2
  }

  let startDay = 0
  if (returnType == 2 || returnType == 11) {
    startDay = 2
  } else if (returnType == 12) {
    startDay = 3
  } else if (returnType == 13) {
    startDay = 4
  } else if (returnType == 14) {
    startDay = 5
  } else if (returnType == 15) {
    startDay = 6
  } else if (returnType == 16) {
    startDay = 7
  } else if (returnType == 17) {
    startDay = 1
  } else {
    return i32.MIN_VALUE
  }

  return ((sundayOne - startDay + 7) % 7) + 1
}

export function excelWeeknumFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE || !isExcelDateSerialInRange(whole)) {
    return i32.MIN_VALUE
  }
  const year = excelYearPartFromSerial(tag, value)
  const month = excelMonthPartFromSerial(tag, value)
  const day = excelDayPartFromSerial(tag, value)
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return i32.MIN_VALUE
  }
  if (returnType == 21) {
    return excelIsoWeeknumValue(whole)
  }

  let weekStartDay = 0
  if (returnType == 1 || returnType == 17) {
    weekStartDay = 0
  } else if (returnType == 2 || returnType == 11) {
    weekStartDay = 1
  } else if (returnType == 12) {
    weekStartDay = 2
  } else if (returnType == 13) {
    weekStartDay = 3
  } else if (returnType == 14) {
    weekStartDay = 4
  } else if (returnType == 15) {
    weekStartDay = 5
  } else if (returnType == 16) {
    weekStartDay = 6
  } else {
    return i32.MIN_VALUE
  }

  const jan1Serial = excelDateSerial(<u8>ValueTag.Number, <f64>year, <u8>ValueTag.Number, 1.0, <u8>ValueTag.Number, 1.0)
  if (isNaN(jan1Serial)) {
    return i32.MIN_VALUE
  }

  const adjustedJan1 = <i32>Math.floor(jan1Serial) - 1
  const jan1Weekday = ((adjustedJan1 % 7) + 7) % 7
  const shift = (jan1Weekday - weekStartDay + 7) % 7

  let dayOfYear = day
  for (let currentMonth = 1; currentMonth < month; currentMonth += 1) {
    dayOfYear += daysInExcelMonth(year, currentMonth)
  }

  return <i32>Math.floor(<f64>(dayOfYear - 1 + shift) / 7.0) + 1
}

export function excelDateSerialFromNumbers(yearNumeric: f64, monthNumeric: f64, dayNumeric: f64): f64 {
  if (!isFinite(yearNumeric) || !isFinite(monthNumeric) || !isFinite(dayNumeric)) {
    return NaN
  }
  let year = <i32>yearNumeric
  const month = <i32>monthNumeric
  const day = <i32>dayNumeric
  if (year >= 0 && year <= 1899) {
    year += 1900
  }
  if (year < 0 || year > 9999) {
    return NaN
  }
  if (year == 1900 && month == 2 && day == 29) {
    return 60
  }

  const zeroBasedMonth = month - 1
  const monthQuotient = floorDiv(zeroBasedMonth, 12)
  const normalizedYear = year + monthQuotient
  const normalizedMonthZero = zeroBasedMonth - monthQuotient * 12
  if (normalizedYear < 0 || normalizedYear > 9999) {
    return NaN
  }
  const days = daysFromCivil(normalizedYear, normalizedMonthZero + 1, 1) + (day - 1)
  let serial = days - EXCEL_EPOCH_DAYS
  if (days >= daysFromCivil(1900, 3, 1)) {
    serial += 1
  }
  return <f64>serial
}

export function excelDateSerial(yearTag: u8, yearValue: f64, monthTag: u8, monthValue: f64, dayTag: u8, dayValue: f64): f64 {
  const year = truncToInt(yearTag, yearValue)
  const month = truncToInt(monthTag, monthValue)
  const day = truncToInt(dayTag, dayValue)
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return NaN
  }
  return excelDateSerialFromNumbers(<f64>year, <f64>month, <f64>day)
}

export function addMonthsExcelSerial(tag: u8, value: f64, offsetTag: u8, offsetValue: f64, endOfMonth: bool): f64 {
  const startYear = excelYearPartFromSerial(tag, value)
  const startMonth = excelMonthPartFromSerial(tag, value)
  const startDay = excelDayPartFromSerial(tag, value)
  const offset = truncToInt(offsetTag, offsetValue)
  if (startYear == i32.MIN_VALUE || startMonth == i32.MIN_VALUE || startDay == i32.MIN_VALUE || offset == i32.MIN_VALUE) {
    return NaN
  }

  const totalMonths = startYear * 12 + (startMonth - 1) + offset
  const shiftedYear = floorDiv(totalMonths, 12)
  const shiftedMonth = totalMonths - shiftedYear * 12 + 1
  if (shiftedYear < 1900 || shiftedYear > 9999) {
    return NaN
  }

  const targetDay = endOfMonth
    ? daysInExcelMonth(shiftedYear, shiftedMonth)
    : min<i32>(startDay, daysInExcelMonth(shiftedYear, shiftedMonth))
  if (shiftedYear == 1900 && shiftedMonth == 2 && targetDay == 29) {
    return 60
  }
  return excelDateSerial(<u8>ValueTag.Number, <f64>shiftedYear, <u8>ValueTag.Number, <f64>shiftedMonth, <u8>ValueTag.Number, <f64>targetDay)
}

export function excelDatedifValue(startWhole: i32, endWhole: i32, unit: string): f64 {
  if (startWhole > endWhole) {
    return NaN
  }

  const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const startDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  const endDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDay == i32.MIN_VALUE
  ) {
    return NaN
  }

  const totalDays = endWhole - startWhole
  const totalMonths = (endYear - startYear) * 12 + (endMonth - startMonth) - (endDay < startDay ? 1 : 0)
  const totalYears = endYear - startYear - (endMonth < startMonth || (endMonth == startMonth && endDay < startDay) ? 1 : 0)

  if (unit == 'D') return <f64>totalDays
  if (unit == 'M') return <f64>totalMonths
  if (unit == 'Y') return <f64>totalYears
  if (unit == 'YM') return <f64>(((totalMonths % 12) + 12) % 12)
  if (unit == 'YD') {
    let comparisonYear = endYear
    let comparison = excelDateSerial(
      <u8>ValueTag.Number,
      <f64>comparisonYear,
      <u8>ValueTag.Number,
      <f64>startMonth,
      <u8>ValueTag.Number,
      <f64>startDay,
    )
    if (isNaN(comparison) || comparison > <f64>endWhole) {
      comparisonYear -= 1
      comparison = excelDateSerial(
        <u8>ValueTag.Number,
        <f64>comparisonYear,
        <u8>ValueTag.Number,
        <f64>startMonth,
        <u8>ValueTag.Number,
        <f64>startDay,
      )
    }
    return isNaN(comparison) ? NaN : <f64>(endWhole - <i32>Math.floor(comparison))
  }
  if (unit == 'MD') {
    if (endDay >= startDay) {
      return <f64>(endDay - startDay)
    }
    const previousMonth = endMonth == 1 ? 12 : endMonth - 1
    return <f64>(daysInExcelMonth(endYear, previousMonth) - startDay + endDay)
  }
  return NaN
}

export function excelDays360Value(startWhole: i32, endWhole: i32, method: i32): f64 {
  const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const startDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
  const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  const endDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDayRaw == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDayRaw == i32.MIN_VALUE
  ) {
    return NaN
  }

  let startDay = startDayRaw
  let endDay = endDayRaw
  if (method == 0) {
    const startIsFebruaryMonthEnd = startMonth == 2 && startDay == daysInExcelMonth(startYear, startMonth)
    const endIsFebruaryMonthEnd = endMonth == 2 && endDay == daysInExcelMonth(endYear, endMonth)
    const startWasDayThirtyOrThirtyOne = startDay >= 30
    if (startDay == 31 || startIsFebruaryMonthEnd) {
      startDay = 30
    }
    if ((endDay == 31 && startWasDayThirtyOrThirtyOne) || (startIsFebruaryMonthEnd && endIsFebruaryMonthEnd)) {
      endDay = 30
    }
  } else {
    if (startDay == 31) {
      startDay = 30
    }
    if (endDay == 31) {
      endDay = 30
    }
  }

  return <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
}

export function excelYearfracValue(startWhole: i32, endWhole: i32, basis: i32): f64 {
  if (basis < 0 || basis > 4) {
    return NaN
  }

  let start = startWhole
  let end = endWhole
  if (start > end) {
    const swapped = start
    start = end
    end = swapped
  }

  let startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>start)
  let startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>start)
  let startDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>start)
  let endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>end)
  let endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>end)
  let endDay = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>end)
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    endYear == i32.MIN_VALUE ||
    endMonth == i32.MIN_VALUE ||
    endDay == i32.MIN_VALUE
  ) {
    return NaN
  }

  let totalDays = 0.0
  if (basis == 0) {
    const startIsFebruaryMonthEnd = startMonth == 2 && startDay == daysInExcelMonth(startYear, startMonth)
    const endIsFebruaryMonthEnd = endMonth == 2 && endDay == daysInExcelMonth(endYear, endMonth)
    const startWasDayThirtyOrThirtyOne = startDay >= 30
    if (startDay == 31 || startIsFebruaryMonthEnd) {
      startDay = 30
    }
    if ((endDay == 31 && startWasDayThirtyOrThirtyOne) || (startIsFebruaryMonthEnd && endIsFebruaryMonthEnd)) {
      endDay = 30
    }
    totalDays = <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
  } else if (basis == 1 || basis == 2 || basis == 3) {
    totalDays = <f64>(end - start)
  } else {
    if (startDay == 31) {
      startDay -= 1
    }
    if (endDay == 31) {
      endDay -= 1
    }
    totalDays = <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
  }

  let daysInYear = 360.0
  if (basis == 1) {
    if (startYear == endYear) {
      daysInYear = isLeapYear(startYear) ? 366.0 : 365.0
    } else {
      const crossesMultipleYears = endYear != startYear + 1 || endMonth < startMonth || (endMonth == startMonth && endDay > startDay)
      if (crossesMultipleYears) {
        let total = 0.0
        for (let year = startYear; year <= endYear; year += 1) {
          total += isLeapYear(year) ? 366.0 : 365.0
        }
        daysInYear = total / <f64>(endYear - startYear + 1)
      } else {
        const startsInLeapYear = isLeapYear(startYear) && (startMonth < 2 || (startMonth == 2 && startDay <= 29))
        const endsInLeapYear = isLeapYear(endYear) && (endMonth > 2 || (endMonth == 2 && endDay == 29))
        daysInYear = startsInLeapYear || endsInLeapYear ? 366.0 : 365.0
      }
    }
  } else if (basis == 3) {
    daysInYear = 365.0
  }

  return totalDays / daysInYear
}

export function securityAnnualizedYearfracValue(settlementWhole: i32, maturityWhole: i32, basis: i32): f64 {
  if (settlementWhole >= maturityWhole) {
    return NaN
  }
  const years = excelYearfracValue(settlementWhole, maturityWhole, basis)
  return isNaN(years) || years <= 0.0 ? NaN : years
}

export function treasuryBillDaysValue(settlementWhole: i32, maturityWhole: i32): f64 {
  if (settlementWhole >= maturityWhole) {
    return NaN
  }
  const days = maturityWhole - settlementWhole
  return days > 0 && days <= 365 ? <f64>days : NaN
}

export function maturityIssueYearfracValue(issueWhole: i32, settlementWhole: i32, maturityWhole: i32, basis: i32): f64 {
  if (issueWhole >= settlementWhole || issueWhole >= maturityWhole || settlementWhole >= maturityWhole) {
    return NaN
  }
  return excelYearfracValue(issueWhole, maturityWhole, basis)
}

export function accruedIssueYearfracValue(issueWhole: i32, settlementWhole: i32, maturityWhole: i32, basis: i32): f64 {
  if (issueWhole >= settlementWhole || issueWhole >= maturityWhole || settlementWhole >= maturityWhole) {
    return NaN
  }
  return excelYearfracValue(issueWhole, settlementWhole, basis)
}

export function couponDaysByBasisValue(startWhole: i32, endWhole: i32, basis: i32): f64 {
  if (basis < 0 || basis > 4 || startWhole > endWhole) {
    return NaN
  }
  if (basis == 0 || basis == 4) {
    const startYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
    const startMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
    const startDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>startWhole)
    const endYear = excelYearPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
    const endMonth = excelMonthPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
    const endDayRaw = excelDayPartFromSerial(<u8>ValueTag.Number, <f64>endWhole)
    if (
      startYear == i32.MIN_VALUE ||
      startMonth == i32.MIN_VALUE ||
      startDayRaw == i32.MIN_VALUE ||
      endYear == i32.MIN_VALUE ||
      endMonth == i32.MIN_VALUE ||
      endDayRaw == i32.MIN_VALUE
    ) {
      return NaN
    }
    let startDay = startDayRaw
    let endDay = endDayRaw
    if (basis == 0) {
      const startIsFebruaryMonthEnd = startMonth == 2 && startDay == daysInExcelMonth(startYear, startMonth)
      const endIsFebruaryMonthEnd = endMonth == 2 && endDay == daysInExcelMonth(endYear, endMonth)
      const startWasDayThirtyOrThirtyOne = startDay >= 30
      if (startDay == 31 || startIsFebruaryMonthEnd) {
        startDay = 30
      }
      if ((endDay == 31 && startWasDayThirtyOrThirtyOne) || (startIsFebruaryMonthEnd && endIsFebruaryMonthEnd)) {
        endDay = 30
      }
    } else {
      if (startDay == 31) {
        startDay = 30
      }
      if (endDay == 31) {
        endDay = 30
      }
    }
    return <f64>((endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay))
  }
  return <f64>(endWhole - startWhole)
}

export function couponDateFromMaturityValue(maturityWhole: i32, periodsBack: i32, frequency: i32): i32 {
  if (periodsBack < 0 || (frequency != 1 && frequency != 2 && frequency != 4)) {
    return i32.MIN_VALUE
  }
  const stepMonths = 12 / frequency
  const serial = addMonthsExcelSerial(<u8>ValueTag.Number, <f64>maturityWhole, <u8>ValueTag.Number, <f64>(-periodsBack * stepMonths), false)
  return isNaN(serial) ? i32.MIN_VALUE : <i32>serial
}

export function couponPeriodsRemainingValue(settlementWhole: i32, maturityWhole: i32, frequency: i32): i32 {
  if (settlementWhole >= maturityWhole || (frequency != 1 && frequency != 2 && frequency != 4)) {
    return i32.MIN_VALUE
  }
  let periodsRemaining = 1
  let previousCoupon = couponDateFromMaturityValue(maturityWhole, periodsRemaining, frequency)
  while (previousCoupon != i32.MIN_VALUE && previousCoupon > settlementWhole) {
    periodsRemaining += 1
    previousCoupon = couponDateFromMaturityValue(maturityWhole, periodsRemaining, frequency)
  }
  return previousCoupon == i32.MIN_VALUE ? i32.MIN_VALUE : periodsRemaining
}

export function couponPeriodDaysValue(previousCoupon: i32, nextCoupon: i32, basis: i32, frequency: i32): f64 {
  if (basis == 1) {
    return couponDaysByBasisValue(previousCoupon, nextCoupon, basis)
  }
  return <f64>(basis == 3 ? 365 : 360) / <f64>frequency
}

export function excelIsoWeeknumValue(whole: i32): i32 {
  if (whole < 0) {
    return i32.MIN_VALUE
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1
  const sundayOne = (((adjustedWhole % 7) + 7) % 7) + 1
  const weekday = sundayOne == 1 ? 7 : sundayOne - 1
  const shiftedWhole = adjustedWhole + 4 - weekday
  const shiftedDays = EXCEL_EPOCH_DAYS + shiftedWhole
  const shiftedYear = civilYear(shiftedDays)
  const yearStart = daysFromCivil(shiftedYear, 1, 1)
  const dayOfYear = shiftedDays - yearStart + 1
  return <i32>Math.floor(<f64>(dayOfYear - 1) / 7.0) + 1
}

function fixedDecliningBalanceRate(cost: f64, salvage: f64, life: f64): f64 {
  if (!isFinite(cost) || !isFinite(salvage) || !isFinite(life) || cost <= 0 || salvage < 0 || life <= 0) {
    return NaN
  }
  const ratio = salvage / cost
  if (ratio < 0) {
    return NaN
  }
  return Math.round((1.0 - Math.pow(ratio, 1.0 / life)) * 1000.0) / 1000.0
}

export function dbDepreciation(cost: f64, salvage: f64, life: f64, period: f64, month: f64): f64 {
  const rate = fixedDecliningBalanceRate(cost, salvage, life)
  const maxPeriod = life + (month < 12.0 ? 1.0 : 0.0)
  if (isNaN(rate) || month < 1 || month > 12 || period < 1 || period > maxPeriod) {
    return NaN
  }

  let bookValue = cost
  let depreciation = 0.0
  for (let currentPeriod = 1.0; currentPeriod <= period; currentPeriod += 1.0) {
    const raw =
      currentPeriod == 1.0
        ? bookValue * rate * (month / 12.0)
        : currentPeriod == <f64>Math.floor(life) + 1.0
          ? bookValue * rate * ((12.0 - month) / 12.0)
          : bookValue * rate
    depreciation = Math.min(Math.max(raw, 0.0), Math.max(0.0, bookValue - salvage))
    bookValue -= depreciation
  }
  return depreciation
}

function ddbPeriodDepreciation(bookValue: f64, salvage: f64, life: f64, factor: f64, remainingLife: f64, noSwitch: bool): f64 {
  const declining = (bookValue * factor) / life
  const straightLine = remainingLife <= 0 ? 0.0 : (bookValue - salvage) / remainingLife
  const base = noSwitch ? declining : Math.max(declining, straightLine)
  return Math.min(Math.max(base, 0.0), Math.max(0.0, bookValue - salvage))
}

export function ddbDepreciation(cost: f64, salvage: f64, life: f64, period: f64, factor: f64): f64 {
  if (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(period) ||
    !isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    period <= 0 ||
    period > life ||
    factor <= 0
  ) {
    return NaN
  }
  let bookValue = cost
  let current = 0.0
  let depreciation = 0.0
  while (current < period && bookValue > salvage) {
    const segment = Math.min(1.0, period - current)
    const full = Math.min(Math.max((bookValue * factor) / life, 0.0), Math.max(0.0, bookValue - salvage))
    depreciation = Math.min(full * segment, Math.max(0.0, bookValue - salvage))
    bookValue -= depreciation
    current += segment
  }
  return depreciation
}

export function vdbDepreciation(cost: f64, salvage: f64, life: f64, startPeriod: f64, endPeriod: f64, factor: f64, noSwitch: bool): f64 {
  if (
    !isFinite(cost) ||
    !isFinite(salvage) ||
    !isFinite(life) ||
    !isFinite(startPeriod) ||
    !isFinite(endPeriod) ||
    !isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    startPeriod < 0 ||
    endPeriod < startPeriod ||
    factor <= 0
  ) {
    return NaN
  }

  let bookValue = cost
  let total = 0.0
  for (let current = 0.0; current < endPeriod && bookValue > salvage; current += 1.0) {
    const overlap = Math.max(0.0, Math.min(endPeriod, current + 1.0) - Math.max(startPeriod, current))
    if (overlap <= 0) {
      const full = ddbPeriodDepreciation(bookValue, salvage, life, factor, life - current, noSwitch)
      bookValue -= full
      continue
    }
    const full = ddbPeriodDepreciation(bookValue, salvage, life, factor, life - current, noSwitch)
    total += full * overlap
    bookValue -= full
  }
  return total
}
