import { ErrorCode, ValueTag, formatGeneralNumberValue, type CellValue } from '@bilig/protocol'

export type GoogleSheetsScalarBuiltin = (...args: CellValue[]) => CellValue

const maxCellTextLength = 32_767

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function stringResult(value: string): CellValue {
  return value.length <= maxCellTextLength ? { tag: ValueTag.String, value, stringId: 0 } : errorValue(ErrorCode.Value)
}

function firstError(args: readonly CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error)
}

function textValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return formatGeneralNumberValue(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return ''
  }
}

function uniqueKey(value: CellValue): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return undefined
    case ValueTag.Number:
      return `n:${formatGeneralNumberValue(value.value)}`
    case ValueTag.Boolean:
      return `b:${value.value ? '1' : '0'}`
    case ValueTag.String:
      return value.value.length === 0 ? undefined : `s:${value.value}`
    case ValueTag.Error:
      return `e:${value.code}`
  }
}

function isEmailText(value: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(value)
}

function isUrlText(value: string): boolean {
  try {
    const parsed = new URL(value)
    return parsed.protocol === 'http:' || parsed.protocol === 'https:'
  } catch {
    return false
  }
}

export const googleSheetsScalarBuiltins: Record<string, GoogleSheetsScalarBuiltin> = {
  COUNTUNIQUE: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    if (args.length === 0) {
      return errorValue(ErrorCode.Value)
    }

    const seen = new Set<string>()
    for (const arg of args) {
      const key = uniqueKey(arg)
      if (key !== undefined) {
        seen.add(key)
      }
    }
    return numberResult(seen.size)
  },
  ISEMAIL: (...args) => {
    const [value] = args
    if (args.length !== 1 || value === undefined) {
      return errorValue(ErrorCode.Value)
    }
    return booleanResult(value.tag === ValueTag.String && isEmailText(value.value))
  },
  ISURL: (...args) => {
    const [value] = args
    if (args.length !== 1 || value === undefined) {
      return errorValue(ErrorCode.Value)
    }
    return booleanResult(value.tag === ValueTag.String && isUrlText(value.value))
  },
  JOIN: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [delimiter, ...values] = args
    if (delimiter === undefined || values.length === 0) {
      return errorValue(ErrorCode.Value)
    }
    return stringResult(values.map(textValue).join(textValue(delimiter)))
  },
  TO_TEXT: (...args) => {
    const existingError = firstError(args)
    if (existingError) {
      return existingError
    }
    const [value] = args
    if (args.length !== 1 || value === undefined) {
      return errorValue(ErrorCode.Value)
    }
    return stringResult(textValue(value))
  },
}
