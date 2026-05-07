import { ValueTag, type CellValue, type ErrorCode } from '@bilig/protocol'

export function emptyValue(): CellValue {
  return { tag: ValueTag.Empty }
}

export function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

export function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

export function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}
