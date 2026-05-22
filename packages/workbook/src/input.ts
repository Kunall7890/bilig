export type WorkbookActionInput =
  | null
  | string
  | number
  | boolean
  | readonly WorkbookActionInput[]
  | { readonly [key: string]: WorkbookActionInput }

export type WorkbookActionInputDescriptionKind = 'json' | 'object' | 'array' | 'string' | 'number' | 'boolean' | 'null'

export interface WorkbookActionInputDescription {
  readonly kind: WorkbookActionInputDescriptionKind
  readonly description?: string
  readonly required?: boolean
  readonly fields?: { readonly [key: string]: WorkbookActionInputDescription }
  readonly items?: WorkbookActionInputDescription
}

export class WorkbookActionInputError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'WorkbookActionInputError'
  }
}

function typeName(value: object): string {
  return value.constructor?.name ?? 'object'
}

function childPath(path: string, key: string): string {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? `${path}.${key}` : `${path}[${JSON.stringify(key)}]`
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function inputDescriptionKind(value: unknown, path: string): WorkbookActionInputDescriptionKind {
  switch (value) {
    case 'json':
    case 'object':
    case 'array':
    case 'string':
    case 'number':
    case 'boolean':
    case 'null':
      return value
    default:
      throw new WorkbookActionInputError(`Action input description at ${path}.kind must be a supported kind`)
  }
}

function normalizeDescriptionText(value: unknown, path: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'string') {
    throw new WorkbookActionInputError(`Action input description at ${path} must be a string`)
  }
  const text = value.trim()
  if (text === '') {
    throw new WorkbookActionInputError(`Action input description at ${path} cannot be empty`)
  }
  return text
}

function normalizeInputDescription(value: unknown, path: string, seen: WeakSet<object>): WorkbookActionInputDescription {
  if (!isPlainObject(value)) {
    throw new WorkbookActionInputError(`Action input description at ${path} must be a plain object`)
  }

  if (seen.has(value)) {
    throw new WorkbookActionInputError(`Action input description at ${path} must not contain cycles`)
  }
  seen.add(value)

  try {
    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw new WorkbookActionInputError(`Action input description at ${path} must not contain symbol keys`)
    }

    const allowedKeys = new Set(['kind', 'description', 'required', 'fields', 'items'])
    const unknownKey = Object.keys(value).find((key) => !allowedKeys.has(key))
    if (unknownKey !== undefined) {
      throw new WorkbookActionInputError(`Action input description at ${childPath(path, unknownKey)} is not supported`)
    }

    const kind = inputDescriptionKind(value['kind'], path)
    const output: {
      kind: WorkbookActionInputDescriptionKind
      description?: string
      required?: boolean
      fields?: { readonly [key: string]: WorkbookActionInputDescription }
      items?: WorkbookActionInputDescription
    } = { kind }

    const description = normalizeDescriptionText(value['description'], `${path}.description`)
    if (description !== undefined) {
      output.description = description
    }

    if (value['required'] !== undefined) {
      if (typeof value['required'] !== 'boolean') {
        throw new WorkbookActionInputError(`Action input description at ${path}.required must be a boolean`)
      }
      output.required = value['required']
    }

    if (value['fields'] !== undefined) {
      if (kind !== 'object') {
        throw new WorkbookActionInputError(`Action input description at ${path}.fields can only be used when kind is object`)
      }
      if (!isPlainObject(value['fields'])) {
        throw new WorkbookActionInputError(`Action input description at ${path}.fields must be a plain object`)
      }
      const fields: Record<string, WorkbookActionInputDescription> = {}
      Object.entries(value['fields'])
        .toSorted(([left], [right]) => left.localeCompare(right))
        .forEach(([key, entry]) => {
          if (key.trim() === '') {
            throw new WorkbookActionInputError(`Action input description at ${path}.fields cannot contain an empty field name`)
          }
          fields[key] = normalizeInputDescription(entry, childPath(`${path}.fields`, key), seen)
        })
      output.fields = Object.freeze(fields)
    }

    if (value['items'] !== undefined) {
      if (kind !== 'array') {
        throw new WorkbookActionInputError(`Action input description at ${path}.items can only be used when kind is array`)
      }
      output.items = normalizeInputDescription(value['items'], `${path}.items`, seen)
    }

    return Object.freeze(output)
  } finally {
    seen.delete(value)
  }
}

export function normalizeWorkbookActionInputDescription(input: unknown): WorkbookActionInputDescription {
  return normalizeInputDescription(input, 'input', new WeakSet())
}

function hasRequiredChild(description: WorkbookActionInputDescription): boolean {
  return (
    description.kind === 'object' &&
    description.fields !== undefined &&
    Object.values(description.fields).some((field) => field.required === true)
  )
}

function inputKind(value: WorkbookActionInput): Exclude<WorkbookActionInputDescriptionKind, 'json'> {
  if (value === null) {
    return 'null'
  }
  if (Array.isArray(value)) {
    return 'array'
  }
  if (typeof value === 'object') {
    return 'object'
  }
  if (typeof value === 'string') {
    return 'string'
  }
  if (typeof value === 'number') {
    return 'number'
  }
  return 'boolean'
}

function isInputObject(value: WorkbookActionInput): value is { readonly [key: string]: WorkbookActionInput } {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateInputAgainstDescription(
  description: WorkbookActionInputDescription,
  value: WorkbookActionInput | undefined,
  path: string,
): void {
  if (value === undefined) {
    if (description.required === true || hasRequiredChild(description)) {
      throw new WorkbookActionInputError(`Action input at ${path} is required`)
    }
    return
  }

  if (description.kind !== 'json' && inputKind(value) !== description.kind) {
    throw new WorkbookActionInputError(`Action input at ${path} must be ${description.kind}`)
  }

  if (description.kind === 'object' && description.fields !== undefined) {
    if (!isInputObject(value)) {
      throw new WorkbookActionInputError(`Action input at ${path} must be object`)
    }
    const unknownKey = Object.keys(value).find((key) => description.fields?.[key] === undefined)
    if (unknownKey !== undefined) {
      throw new WorkbookActionInputError(`Action input at ${childPath(path, unknownKey)} is not supported`)
    }
    Object.entries(description.fields).forEach(([key, field]) => {
      const childValue = Object.hasOwn(value, key) ? value[key] : undefined
      validateInputAgainstDescription(field, childValue, childPath(path, key))
    })
    return
  }

  if (description.kind === 'array' && description.items !== undefined) {
    if (!Array.isArray(value)) {
      throw new WorkbookActionInputError(`Action input at ${path} must be array`)
    }
    value.forEach((entry, index) => {
      validateInputAgainstDescription(description.items!, entry, `${path}[${index.toString()}]`)
    })
  }
}

export function validateWorkbookActionInput(description: WorkbookActionInputDescription, input: WorkbookActionInput | undefined): void {
  validateInputAgainstDescription(description, input, 'input')
}

function normalizeInput(value: unknown, path: string, seen: WeakSet<object>): WorkbookActionInput {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new WorkbookActionInputError(`Action input at ${path} must be a finite number`)
    }
    return value
  }

  if (value === undefined) {
    throw new WorkbookActionInputError(`Action input at ${path} must not be undefined`)
  }

  if (typeof value === 'bigint' || typeof value === 'function' || typeof value === 'symbol') {
    throw new WorkbookActionInputError(`Action input at ${path} must be JSON-safe, not ${typeof value}`)
  }

  if (typeof value !== 'object') {
    throw new WorkbookActionInputError(`Action input at ${path} must be JSON-safe`)
  }

  if (seen.has(value)) {
    throw new WorkbookActionInputError(`Action input at ${path} must not contain cycles`)
  }
  seen.add(value)

  try {
    if (Array.isArray(value)) {
      const output: WorkbookActionInput[] = []
      for (let index = 0; index < value.length; index += 1) {
        if (!Object.hasOwn(value, index)) {
          throw new WorkbookActionInputError(`Action input at ${path}[${index}] must not be a sparse array hole`)
        }
        output.push(normalizeInput(value[index], `${path}[${index}]`, seen))
      }
      return output
    }

    if (!isPlainObject(value)) {
      throw new WorkbookActionInputError(`Action input at ${path} must be a plain JSON object, not ${typeName(value)}`)
    }

    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw new WorkbookActionInputError(`Action input at ${path} must not contain symbol keys`)
    }

    const output: { [key: string]: WorkbookActionInput } = {}
    Object.entries(value)
      .toSorted(([left], [right]) => left.localeCompare(right))
      .forEach(([key, entry]) => {
        output[key] = normalizeInput(entry, childPath(path, key), seen)
      })
    return output
  } finally {
    seen.delete(value)
  }
}

export function normalizeWorkbookActionInput(input: unknown): WorkbookActionInput {
  return normalizeInput(input, 'input', new WeakSet())
}

export function normalizeOptionalWorkbookActionInput(input: WorkbookActionInput | undefined): WorkbookActionInput | undefined {
  return input === undefined ? undefined : normalizeWorkbookActionInput(input)
}

export function hasOwnActionInput(value: object): boolean {
  return Object.prototype.hasOwnProperty.call(value, 'input')
}

export function getOwnActionInput(value: object): unknown {
  return Object.getOwnPropertyDescriptor(value, 'input')?.value
}
