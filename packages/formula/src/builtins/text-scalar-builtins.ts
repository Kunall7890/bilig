import { ErrorCode, type CellValue } from '@bilig/protocol'
import type { TextBuiltin } from './text.js'

interface TextScalarBuiltinDeps {
  error: (code: ErrorCode) => CellValue
  stringResult: (value: string) => CellValue
  numberResult: (value: number) => CellValue
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined
  coerceText: (value: CellValue) => string
  coerceNumber: (value: CellValue) => number | undefined
  isErrorValue: (value: number | CellValue) => value is CellValue
}

function charCodeFromArgument(deps: TextScalarBuiltinDeps, value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return deps.error(ErrorCode.Value)
  }
  const code = deps.coerceNumber(value)
  if (code === undefined) {
    return deps.error(ErrorCode.Value)
  }
  const integerCode = Math.trunc(code)
  if (!Number.isFinite(integerCode) || integerCode < 1 || integerCode > 255) {
    return deps.error(ErrorCode.Value)
  }
  return integerCode
}

function unicodeCodeFromArgument(deps: TextScalarBuiltinDeps, value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return deps.error(ErrorCode.Value)
  }
  const code = deps.coerceNumber(value)
  if (code === undefined) {
    return deps.error(ErrorCode.Value)
  }
  const integerCode = Math.trunc(code)
  if (!Number.isFinite(integerCode) || integerCode < 1 || integerCode > 0x10ffff) {
    return deps.error(ErrorCode.Value)
  }
  if (integerCode >= 0xd800 && integerCode <= 0xdfff) {
    return deps.error(ErrorCode.NA)
  }
  return integerCode
}

function firstCodePoint(deps: TextScalarBuiltinDeps, value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return deps.error(ErrorCode.Value)
  }
  const text = deps.coerceText(value)
  if (text.length === 0) {
    return deps.error(ErrorCode.Value)
  }
  const codePoint = text.codePointAt(0)
  return codePoint === undefined ? deps.error(ErrorCode.Value) : codePoint
}

export function createTextScalarBuiltins(deps: TextScalarBuiltinDeps): Record<string, TextBuiltin> {
  return {
    CHAR: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [codeValue] = args
      const codePoint = charCodeFromArgument(deps, codeValue)
      if (deps.isErrorValue(codePoint)) {
        return codePoint
      }
      return deps.stringResult(String.fromCodePoint(codePoint))
    },
    CODE: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue] = args
      const codePoint = firstCodePoint(deps, textValue)
      return deps.isErrorValue(codePoint) ? codePoint : deps.numberResult(codePoint)
    },
    UNICODE: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [textValue] = args
      const codePoint = firstCodePoint(deps, textValue)
      return deps.isErrorValue(codePoint) ? codePoint : deps.numberResult(codePoint)
    },
    UNICHAR: (...args) => {
      const existingError = deps.firstError(args)
      if (existingError) {
        return existingError
      }
      const [codeValue] = args
      const codePoint = unicodeCodeFromArgument(deps, codeValue)
      if (deps.isErrorValue(codePoint)) {
        return codePoint
      }
      return deps.stringResult(String.fromCodePoint(codePoint))
    },
  }
}
