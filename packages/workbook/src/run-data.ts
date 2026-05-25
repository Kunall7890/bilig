export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    return Array.from({ length: value.length }, (_, index) => {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        return undefined
      }
      if (!('value' in descriptor)) {
        throw new Error('Accessor values cannot be canonicalized')
      }
      return canonicalValue(descriptor.value)
    })
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(Object.getOwnPropertyDescriptors(value))
        .filter(([, descriptor]) => descriptor.enumerable)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, descriptor]) => {
          if (!('value' in descriptor)) {
            throw new Error('Accessor values cannot be canonicalized')
          }
          return [key, canonicalValue(descriptor.value)]
        }),
    )
  }
  return value
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

export function ownDataDescriptor(value: object, key: string): PropertyDescriptor | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor : undefined
}

export function ownValue(value: object, key: string): unknown {
  return ownDataDescriptor(value, key)?.value
}

export function arrayDataValues<T>(value: unknown, guard: (entry: unknown) => entry is T): readonly T[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  if (firstAccessorPath(value, 'array') !== null) {
    return null
  }

  const entries: T[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !guard(descriptor.value)) {
      return null
    }
    entries.push(descriptor.value)
  }
  return entries
}

export function cloneData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
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
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        continue
      }
      if (!('value' in descriptor)) {
        throw new Error('Transport data must not contain accessors')
      }
      cloned[index] = cloneData(descriptor.value, seen)
    }
    return cloned
  }
  const cloned: Record<string, unknown> = {}
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable) {
      return
    }
    if (!('value' in descriptor)) {
      throw new Error('Transport data must not contain accessors')
    }
    Object.defineProperty(cloned, key, {
      configurable: true,
      enumerable: true,
      value: cloneData(descriptor.value, seen),
      writable: true,
    })
  })
  return cloned
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
