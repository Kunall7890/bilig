import type { LiteralInput } from '@bilig/protocol'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function stringAttribute(value: unknown): string | null {
  return typeof value === 'string' && value.trim().length > 0 ? value : null
}

function readSharedItemValue(entry: unknown): LiteralInput | undefined {
  if (!isRecord(entry)) {
    return undefined
  }
  const value = stringAttribute(entry['v'])
  if (value === null) {
    return null
  }
  return value
}

function readSharedNumericItem(entry: unknown): LiteralInput | undefined {
  if (!isRecord(entry)) {
    return undefined
  }
  const raw = stringAttribute(entry['v'])
  if (raw === null) {
    return undefined
  }
  const value = Number(raw)
  return Number.isFinite(value) ? value : raw
}

function readSharedBooleanItem(entry: unknown): LiteralInput | undefined {
  if (!isRecord(entry)) {
    return undefined
  }
  const raw = stringAttribute(entry['v'])
  return raw === null ? undefined : raw === '1' || raw.toLowerCase() === 'true'
}

export function readCacheFieldSharedItems(cacheField: Record<string, unknown>): LiteralInput[] {
  const sharedItems = recordChild(cacheField, 'sharedItems')
  if (!sharedItems) {
    return []
  }
  const values: LiteralInput[] = []
  for (const item of asArray(sharedItems['s'])) {
    const value = readSharedItemValue(item)
    if (value !== undefined) {
      values.push(value)
    }
  }
  for (const item of asArray(sharedItems['n'])) {
    const value = readSharedNumericItem(item)
    if (value !== undefined) {
      values.push(value)
    }
  }
  for (const item of asArray(sharedItems['b'])) {
    const value = readSharedBooleanItem(item)
    if (value !== undefined) {
      values.push(value)
    }
  }
  for (const item of asArray(sharedItems['e'])) {
    const value = readSharedItemValue(item)
    if (value !== undefined) {
      values.push(value)
    }
  }
  values.push(...asArray(sharedItems['m']).map(() => null))
  return values
}
