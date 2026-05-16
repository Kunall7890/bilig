export function resolveBrowserTestCiMode(value: string | undefined): boolean {
  if (value === undefined || value.length === 0 || value === '0' || value === 'false') {
    return false
  }
  if (value === '1' || value === 'true') {
    return true
  }
  throw new Error(`CI must be "1", "true", "0", or "false" when set, got ${value}`)
}

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
