const simpleNumericTextPattern = /^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:[eE][+-]?\d+)?$/

function trimAsciiWhitespace(value: string): string {
  let start = 0
  let end = value.length
  while (start < end && value.charCodeAt(start) <= 32) {
    start += 1
  }
  while (end > start && value.charCodeAt(end - 1) <= 32) {
    end -= 1
  }
  return value.slice(start, end)
}

export function parseNumericText(value: string): number | undefined {
  const trimmed = trimAsciiWhitespace(value)
  if (trimmed.length === 0) {
    return undefined
  }

  if (simpleNumericTextPattern.test(trimmed)) {
    const direct = Number(trimmed)
    if (Number.isFinite(direct)) {
      return direct
    }
  }

  if (!trimmed.includes(',')) {
    return undefined
  }

  const grouped = /^([+-]?)(\d{1,3}(?:,\d{3})+)(\.\d*)?([eE][+-]?\d+)?$/.exec(trimmed)
  if (!grouped) {
    return undefined
  }

  const normalized = `${grouped[1] ?? ''}${(grouped[2] ?? '').replaceAll(',', '')}${grouped[3] ?? ''}${grouped[4] ?? ''}`
  const numeric = Number(normalized)
  return Number.isFinite(numeric) ? numeric : undefined
}

export function parseArithmeticNumericText(value: string): number | undefined {
  return value === '' ? 0 : parseNumericText(value)
}
