export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

export function stringValue(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

export function numberValue(value: unknown): number | null {
  if (typeof value !== 'string' || value.trim().length === 0) {
    return null
  }
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

export function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

export function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return `#${normalized.toLowerCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/.test(normalized)) {
    return `#${normalized.slice(2).toLowerCase()}`
  }
  return null
}

export function toArgbColor(value: string): string | null {
  const normalized = normalizeRgbColor(value)
  return normalized ? `FF${normalized.slice(1).toUpperCase()}` : null
}
