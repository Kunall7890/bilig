import type { CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

function isStrictBuiltinArity(name: string, argc: number): boolean {
  switch (name) {
    case 'ABS':
    case 'ACOS':
    case 'ACOSH':
    case 'ACOT':
    case 'ACOTH':
    case 'ASIN':
    case 'ASINH':
    case 'ATAN':
    case 'ATANH':
    case 'BIN2DEC':
    case 'CHAR':
    case 'CLEAN':
    case 'CODE':
    case 'COS':
    case 'COSH':
    case 'COT':
    case 'COTH':
    case 'CSC':
    case 'CSCH':
    case 'DAY':
    case 'DEGREES':
    case 'EVEN':
    case 'EXP':
    case 'FACT':
    case 'FACTDOUBLE':
    case 'HEX2DEC':
    case 'HOUR':
    case 'INT':
    case 'LEN':
    case 'LENB':
    case 'LN':
    case 'LOG10':
    case 'LOWER':
    case 'MINUTE':
    case 'MONTH':
    case 'ODD':
    case 'OCT2DEC':
    case 'RADIANS':
    case 'SECOND':
    case 'SEC':
    case 'SECH':
    case 'SIGN':
    case 'SIN':
    case 'SINH':
    case 'SQRT':
    case 'TAN':
    case 'TANH':
    case 'TIMEVALUE':
    case 'TRIM':
    case 'UNICODE':
    case 'UNICHAR':
    case 'UPPER':
    case 'VALUE':
    case 'YEAR':
      return argc === 1
    case 'BIN2HEX':
    case 'BIN2OCT':
    case 'DEC2BIN':
    case 'DEC2HEX':
    case 'DEC2OCT':
    case 'DELTA':
    case 'GESTEP':
    case 'HEX2BIN':
    case 'HEX2OCT':
    case 'LEFT':
    case 'LOG':
    case 'OCT2BIN':
    case 'OCT2HEX':
    case 'RIGHT':
    case 'ROUND':
    case 'ROUNDDOWN':
    case 'ROUNDUP':
      return argc === 1 || argc === 2
    case 'BASE':
    case 'DAYS360':
    case 'FIND':
    case 'SEARCH':
    case 'YEARFRAC':
      return argc === 2 || argc === 3
    case 'ADDRESS':
      return argc >= 2 && argc <= 5
    case 'COMBIN':
    case 'COMBINA':
    case 'DAYS':
    case 'DECIMAL':
    case 'DOLLARDE':
    case 'DOLLARFR':
    case 'EDATE':
    case 'EOMONTH':
    case 'EXACT':
    case 'POWER':
    case 'QUOTIENT':
    case 'REPT':
    case 'TEXT':
      return argc === 2
    case 'CONVERT':
    case 'DATE':
    case 'MID':
    case 'TIME':
      return argc === 3
    case 'REPLACE':
      return argc === 4
    case 'SUBSTITUTE':
      return argc === 3 || argc === 4
    default:
      return true
  }
}

export function isFormulaCallArity(name: string, argc: number): boolean {
  if (!isStrictBuiltinArity(name, argc)) {
    return false
  }
  switch (name) {
    case 'WORKDAY':
    case 'NETWORKDAYS':
      return argc === 2 || argc === 3
    case 'WORKDAY.INTL':
    case 'NETWORKDAYS.INTL':
      return argc >= 2 && argc <= 4
    default:
      return true
  }
}

export function enforceBuiltinArities(map: Record<string, Builtin>, valueError: () => EvaluationResult): Record<string, Builtin> {
  return Object.fromEntries(
    Object.entries(map).map(([name, builtin]) => [
      name,
      (...args: CellValue[]): EvaluationResult => (isStrictBuiltinArity(name, args.length) ? builtin(...args) : valueError()),
    ]),
  )
}
