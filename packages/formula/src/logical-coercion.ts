import { ErrorCode, ValueTag, type CellValue, type ErrorValue } from '@bilig/protocol'

export type LogicalCoercion = { ok: true; value: boolean } | { ok: false; error: ErrorValue }

function errorValue(code: ErrorCode): ErrorValue {
  return { tag: ValueTag.Error, code }
}

export function coerceLogicalText(value: string): LogicalCoercion {
  if (value === '') {
    return { ok: true, value: false }
  }
  const normalized = value.toUpperCase()
  if (normalized === 'TRUE') {
    return { ok: true, value: true }
  }
  if (normalized === 'FALSE') {
    return { ok: true, value: false }
  }
  return { ok: false, error: errorValue(ErrorCode.Value) }
}

export function coerceLogicalValue(value: CellValue): LogicalCoercion {
  switch (value.tag) {
    case ValueTag.Boolean:
      return { ok: true, value: value.value }
    case ValueTag.Number:
      return { ok: true, value: value.value !== 0 }
    case ValueTag.Empty:
      return { ok: true, value: false }
    case ValueTag.String:
      return coerceLogicalText(value.value)
    case ValueTag.Error:
      return { ok: false, error: value }
  }
}
