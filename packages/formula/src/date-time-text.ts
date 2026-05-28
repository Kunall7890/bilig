import {
  daysInExcelMonth,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  type ExcelDateParts,
  type ExcelDateSystem,
} from './builtins/excel-date.js'

const SECONDS_PER_DAY = 86_400

export function parseDateValueFromText(raw: string, dateSystem: ExcelDateSystem): number | undefined {
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
    const year = normalizeDateTextYear(thirdNumber, third.length)
    return firstNumber > 12 && secondNumber >= 1 && secondNumber <= 12
      ? { year, month: secondNumber, day: firstNumber }
      : { year, month: firstNumber, day: secondNumber }
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

interface ParsedTimeTextSerial {
  elapsedDays: number
  timeFraction: number
}

function parseTimeTextSerial(raw: string): ParsedTimeTextSerial | undefined {
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

  if (truncHours < 0 || truncMinutes < 0 || truncSeconds < 0 || truncHours > 32_767 || truncMinutes > 32_767 || truncSeconds > 32_767) {
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
  }

  const totalSeconds = hourValue * 3600 + truncMinutes * 60 + truncSeconds
  const secondsOfDay = ((totalSeconds % SECONDS_PER_DAY) + SECONDS_PER_DAY) % SECONDS_PER_DAY
  return {
    elapsedDays: totalSeconds / SECONDS_PER_DAY,
    timeFraction: secondsOfDay / SECONDS_PER_DAY,
  }
}

export function parseTimeValueText(raw: string): number | undefined {
  return parseTimeTextSerial(raw)?.timeFraction
}

export function parseElapsedTimeValueText(raw: string): number | undefined {
  return parseTimeTextSerial(raw)?.elapsedDays
}
