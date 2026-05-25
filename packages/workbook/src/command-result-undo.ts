import { isWorkbookOp } from './guards.js'
import type { EngineOp } from './ops.js'
import type { WorkbookUndoRef } from './result.js'

export function normalizeCommandResultUndoRef(value: unknown, path: string): WorkbookUndoRef {
  if (!isRecord(value)) {
    throw new Error(`Workbook command result ${path} must be an object`)
  }
  const accessorKeys = ownAccessorKeys(value, ['id', 'ops'])
  if (accessorKeys.length > 0) {
    throw new Error(`Workbook command result ${path}.${accessorKeys[0]} must be a data property`)
  }
  const id = ownValue(value, 'id')
  if (typeof id !== 'string') {
    throw new Error(`Workbook command result ${path}.id must be a string`)
  }
  const ops = ownValue(value, 'ops')
  if (ops !== undefined && !Array.isArray(ops)) {
    throw new Error(`Workbook command result ${path}.ops must be an array`)
  }
  if (Array.isArray(ops)) {
    const accessorPath = firstAccessorPath(ops, `${path}.ops`)
    if (accessorPath !== null) {
      throw new Error(`Workbook command result ${accessorPath} must contain only data properties`)
    }
    const normalizedOps: EngineOp[] = []
    for (let index = 0; index < ops.length; index += 1) {
      const op = arrayDataValue(ops, index)
      if (op === undefined) {
        throw new Error(`Workbook command result ${path}.ops[${String(index)}] must contain only data properties`)
      }
      normalizedOps.push(normalizeOp(op))
    }
    return Object.freeze({
      id: normalizeExactString(id, `${path}.id`),
      ops: Object.freeze(normalizedOps),
    })
  }
  return Object.freeze({
    id: normalizeExactString(id, `${path}.id`),
  })
}

function normalizeOp(value: unknown): EngineOp {
  if (!isWorkbookOp(value)) {
    throw new Error('Workbook command result op is invalid')
  }
  const cloned = cloneData(value)
  if (!isWorkbookOp(cloned)) {
    throw new Error('Workbook command result op clone is invalid')
  }
  return freezeData(cloned)
}

function normalizeExactString(value: string, path: string): string {
  const normalized = value.trim()
  if (normalized === '') {
    throw new Error(`Workbook command result ${path} cannot be empty`)
  }
  if (normalized !== value) {
    throw new Error(`Workbook command result ${path} must not have leading or trailing whitespace`)
  }
  return normalized
}

function arrayDataValue(value: readonly unknown[], index: number): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
  return descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
}

function ownValue(value: Record<string, unknown>, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function ownAccessorKeys(value: Record<string, unknown>, keys: readonly string[]): readonly string[] {
  const accessors: string[] = []
  for (const key of keys) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    if (descriptor !== undefined && !('value' in descriptor)) {
      accessors.push(key)
    }
  }
  return accessors
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function firstAccessorPath(value: unknown, path: string, seen = new WeakSet<object>()): string | null {
  if (typeof value !== 'object' || value === null) {
    return null
  }
  if (seen.has(value)) {
    return null
  }
  seen.add(value)

  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    const childPath = Array.isArray(value) && /^\d+$/.test(key) ? `${path}[${key}]` : `${path}.${key}`
    if (!('value' in descriptor)) {
      return childPath
    }
    const nestedPath = firstAccessorPath(descriptor.value, childPath, seen)
    if (nestedPath !== null) {
      return nestedPath
    }
  }
  return null
}

function freezeData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function cloneData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  const existing = seen.get(value)
  if (existing !== undefined) {
    return existing
  }
  if (Array.isArray(value)) {
    const cloned: unknown[] = []
    seen.set(value, cloned)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
      if (descriptor !== undefined && descriptor.enumerable && 'value' in descriptor) {
        cloned[index] = cloneData(descriptor.value, seen)
      }
    }
    return cloned
  }
  const cloned: Record<string, unknown> = {}
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (descriptor.enumerable && 'value' in descriptor) {
      Object.defineProperty(cloned, key, {
        configurable: true,
        enumerable: true,
        value: cloneData(descriptor.value, seen),
        writable: true,
      })
    }
  })
  return cloned
}
