export function parsePositiveIntegerEnv(value: string | undefined, fallback: number, name: string): number {
  if (value === undefined || value.length === 0) {
    return fallback
  }

  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`${name} must be a positive integer, got ${value}`)
  }

  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`${name} must be a safe integer, got ${value}`)
  }

  return parsed
}
