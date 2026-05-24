export type OptionalDataValue =
  | {
      readonly status: 'missing'
    }
  | {
      readonly status: 'present'
      readonly value: unknown
    }

export function isObject(value: unknown): value is object {
  return typeof value === 'object' && value !== null
}

export function isObjectRecord(value: unknown): value is object {
  if (!isObject(value) || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

export function optionalDataProperty(value: object, key: string, label: string): OptionalDataValue {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return { status: 'missing' }
  }
  if (!('value' in descriptor)) {
    throw new Error(`${label} must be a data property`)
  }
  return {
    status: 'present',
    value: descriptor.value,
  }
}

export function requiredDataProperty(value: object, key: string, label: string): unknown {
  const property = optionalDataProperty(value, key, label)
  if (property.status === 'missing') {
    throw new Error(`${label} must be a data property`)
  }
  return property.value
}
