import { readFileSync } from 'node:fs'

export function readJsonObject(path: string): Record<string, unknown> {
  return asObject(JSON.parse(readFileSync(path, 'utf8')) as unknown, path)
}

export function objectField(value: Record<string, unknown>, field: string): Record<string, unknown> {
  return asObject(value[field], field)
}

export function stringField(value: Record<string, unknown>, field: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${field} to be a string`)
  }
  return fieldValue
}

export function optionalStringField(value: Record<string, unknown>, field: string): string | null {
  const fieldValue = value[field]
  if (fieldValue === null) {
    return null
  }
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${field} to be a string or null`)
  }
  return fieldValue
}

export function stringArrayField(value: Record<string, unknown>, field: string): string[] {
  const fieldValue = arrayField(value, field)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${field} to be a string array`)
  }
  return fieldValue
}

export function numberArrayField(value: Record<string, unknown>, field: string): number[] {
  const fieldValue = arrayField(value, field)
  if (!fieldValue.every(isFiniteNumber)) {
    throw new Error(`Expected ${field} to be a finite number array`)
  }
  return fieldValue
}

export function arrayField(value: Record<string, unknown>, field: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${field} to be an array`)
  }
  return fieldValue
}

export function numberField(value: Record<string, unknown>, field: string): number {
  const fieldValue = value[field]
  if (!isFiniteNumber(fieldValue)) {
    throw new Error(`Expected ${field} to be a finite number`)
  }
  return fieldValue
}

export function optionalNumberField(value: Record<string, unknown>, field: string): number | null {
  const fieldValue = value[field]
  if (fieldValue === null) {
    return null
  }
  if (!isFiniteNumber(fieldValue)) {
    throw new Error(`Expected ${field} to be a finite number or null`)
  }
  return fieldValue
}

export function booleanField(value: Record<string, unknown>, field: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${field} to be a boolean`)
  }
  return fieldValue
}

export function literalField<const T extends string | number | boolean>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${String(expected)}`)
  }
  return expected
}

export function asObject(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

export function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}
