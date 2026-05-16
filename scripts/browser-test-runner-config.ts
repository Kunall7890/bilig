export function resolveBrowserTestTimeoutMs(value: string | undefined, fallbackMs: number): number {
  if (value === undefined || value.length === 0) {
    return fallbackMs
  }

  const parsed = parsePositiveDecimalInteger(value)
  return parsed ?? fallbackMs
}

export function resolveBrowserTestConfiguredPort(value: string, label: string): string {
  const parsed = parsePositiveDecimalInteger(value)
  if (parsed === undefined || parsed > 65_535) {
    throw new Error(`${label} must be a TCP port between 1 and 65535, got ${value}`)
  }
  return String(parsed)
}

function parsePositiveDecimalInteger(value: string): number | undefined {
  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    return undefined
  }
  const parsed = Number(value)
  return Number.isSafeInteger(parsed) ? parsed : undefined
}
