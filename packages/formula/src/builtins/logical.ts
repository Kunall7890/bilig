import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createBlockedBuiltinMap, logicalPlaceholderBuiltinNames } from './placeholder.js'
import { parseArithmeticNumericText } from '../numeric-text.js'
import { coerceLogicalValue } from '../logical-coercion.js'

export type LogicalBuiltin = (...args: CellValue[]) => CellValue

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function requireOneArg(args: CellValue[], evaluate: (value: CellValue) => CellValue): CellValue {
  return args.length === 1 ? evaluate(args[0]!) : errorValue(ErrorCode.Value)
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value }
}

function isError(value: CellValue): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value.tag === ValueTag.Error
}

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase()
  const normalizedRight = right.toUpperCase()
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  return normalizedLeft < normalizedRight ? -1 : 1
}

function compareSwitchScalars(left: CellValue, right: CellValue): number | undefined {
  const leftTextLike = left.tag === ValueTag.String || left.tag === ValueTag.Empty
  const rightTextLike = right.tag === ValueTag.String || right.tag === ValueTag.Empty
  if (leftTextLike && rightTextLike) {
    return compareText(left.tag === ValueTag.String ? left.value : '', right.tag === ValueTag.String ? right.value : '')
  }
  if (left.tag === ValueTag.Number && right.tag === ValueTag.Number) {
    if (left.value === right.value) {
      return 0
    }
    return left.value < right.value ? -1 : 1
  }
  if (left.tag === ValueTag.Boolean && right.tag === ValueTag.Boolean) {
    if (left.value === right.value) {
      return 0
    }
    return left.value ? 1 : -1
  }
  return undefined
}

function coerceNumberLike(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
      return parseArithmeticNumericText(value.value)
    case ValueTag.Error:
    default:
      return undefined
  }
}

function errorTypeCode(code: ErrorCode): number | undefined {
  switch (code) {
    case ErrorCode.Null:
      return 1
    case ErrorCode.Div0:
      return 2
    case ErrorCode.Value:
      return 3
    case ErrorCode.Ref:
      return 4
    case ErrorCode.Name:
      return 5
    case ErrorCode.Num:
      return 6
    case ErrorCode.NA:
      return 7
    case ErrorCode.None:
    case ErrorCode.Cycle:
    case ErrorCode.Spill:
    case ErrorCode.Blocked:
    case ErrorCode.Field:
    default:
      return undefined
  }
}

const logicalPlaceholderBuiltins = createBlockedBuiltinMap(logicalPlaceholderBuiltinNames)

export const logicalBuiltins: Record<string, LogicalBuiltin> = {
  TRUE: (...args) => (args.length === 0 ? booleanResult(true) : errorValue(ErrorCode.Value)),
  FALSE: (...args) => (args.length === 0 ? booleanResult(false) : errorValue(ErrorCode.Value)),
  NA: (...args) => {
    if (args.length > 0) {
      return errorValue(ErrorCode.Value)
    }
    return errorValue(ErrorCode.NA)
  },
  IF: (...args) => {
    const [condition, truthy, falsy] = args
    if (condition === undefined || truthy === undefined) {
      return errorValue(ErrorCode.Value)
    }

    const coerced = coerceLogicalValue(condition)
    if (!coerced.ok) {
      return coerced.error
    }

    return coerced.value ? truthy : args.length >= 3 ? (falsy ?? emptyValue()) : booleanResult(false)
  },
  IFERROR: (value, valueIfError = emptyValue()) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value)
    }
    return isError(value) ? valueIfError : value
  },
  IFNA: (value, valueIfNa = emptyValue()) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value)
    }
    return isError(value) && value.code === ErrorCode.NA ? valueIfNa : value
  },
  AND: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value)
    }

    let firstError: Extract<CellValue, { tag: ValueTag.Error }> | undefined
    let result = true
    for (const arg of args) {
      const coerced = coerceLogicalValue(arg)
      if (!coerced.ok) {
        firstError ??= coerced.error
        continue
      }
      if (!coerced.value) {
        result = false
      }
    }

    return firstError ?? booleanResult(result)
  },
  OR: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value)
    }

    let firstError: Extract<CellValue, { tag: ValueTag.Error }> | undefined
    let result = false
    for (const arg of args) {
      const coerced = coerceLogicalValue(arg)
      if (!coerced.ok) {
        firstError ??= coerced.error
        continue
      }
      if (coerced.value) {
        result = true
      }
    }

    return firstError ?? booleanResult(result)
  },
  NOT: (value) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value)
    }

    const coerced = coerceLogicalValue(value)
    if (!coerced.ok) {
      return coerced.error
    }

    return booleanResult(!coerced.value)
  },
  XOR: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value)
    }

    let parity = false
    for (const arg of args) {
      const coerced = coerceLogicalValue(arg)
      if (!coerced.ok) {
        return coerced.error
      }
      parity = parity !== coerced.value
    }

    return booleanResult(parity)
  },
  IFS: (...args) => {
    if (args.length < 2 || args.length % 2 !== 0) {
      return errorValue(ErrorCode.Value)
    }

    for (let index = 0; index < args.length; index += 2) {
      const condition = args[index]!
      const result = args[index + 1]!
      const coerced = coerceLogicalValue(condition)
      if (!coerced.ok) {
        return coerced.error
      }
      if (coerced.value) {
        return result
      }
    }

    return errorValue(ErrorCode.NA)
  },
  SWITCH: (...args) => {
    if (args.length < 3) {
      return errorValue(ErrorCode.Value)
    }

    const expression = args[0]!
    if (isError(expression)) {
      return expression
    }

    const hasDefault = (args.length - 1) % 2 === 1
    const pairLimit = hasDefault ? args.length - 1 : args.length
    for (let index = 1; index < pairLimit; index += 2) {
      const candidate = args[index]!
      if (isError(candidate)) {
        return candidate
      }
      if (compareSwitchScalars(expression, candidate) === 0) {
        return args[index + 1]!
      }
    }

    return hasDefault ? args[args.length - 1]! : errorValue(ErrorCode.NA)
  },
  ISBLANK: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.Empty)),
  ISNUMBER: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.Number)),
  ISTEXT: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.String)),
  ISERROR: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.Error)),
  ISERR: (...args) =>
    requireOneArg(args, (value) => {
      if (value.tag !== ValueTag.Error) {
        return booleanResult(false)
      }
      return booleanResult(value.code !== ErrorCode.NA)
    }),
  ISFORMULA: (...args) => requireOneArg(args, () => booleanResult(false)),
  ISLOGICAL: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.Boolean)),
  ISNONTEXT: (...args) => requireOneArg(args, (value) => booleanResult(value.tag !== ValueTag.String)),
  ISEVEN: (...args) =>
    requireOneArg(args, (value) => {
      if (value.tag === ValueTag.Error) {
        return value
      }
      const numberValue = coerceNumberLike(value)
      if (numberValue === undefined) {
        return errorValue(ErrorCode.Value)
      }
      return booleanResult(Math.trunc(numberValue) % 2 === 0)
    }),
  ISODD: (...args) =>
    requireOneArg(args, (value) => {
      if (value.tag === ValueTag.Error) {
        return value
      }
      const numberValue = coerceNumberLike(value)
      if (numberValue === undefined) {
        return errorValue(ErrorCode.Value)
      }
      return booleanResult(Math.trunc(numberValue) % 2 !== 0)
    }),
  ISNA: (...args) => requireOneArg(args, (value) => booleanResult(value.tag === ValueTag.Error && value.code === ErrorCode.NA)),
  ISREF: (...args) => requireOneArg(args, () => booleanResult(false)),
  'ERROR.TYPE': (...args) =>
    requireOneArg(args, (value) => {
      if (value.tag !== ValueTag.Error) {
        return errorValue(ErrorCode.NA)
      }
      const code = errorTypeCode(value.code)
      return code === undefined ? errorValue(ErrorCode.NA) : { tag: ValueTag.Number, value: code }
    }),
  ...logicalPlaceholderBuiltins,
}

export function getLogicalBuiltin(name: string): LogicalBuiltin | undefined {
  return logicalBuiltins[name.toUpperCase()]
}
