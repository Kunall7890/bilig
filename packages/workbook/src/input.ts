export type WorkbookActionInput =
  | null
  | string
  | number
  | boolean
  | readonly WorkbookActionInput[]
  | { readonly [key: string]: WorkbookActionInput }

export type WorkbookActionInputDescriptionKind = 'json' | 'object' | 'array' | 'string' | 'number' | 'boolean' | 'null'

export const workbookActionInputDescriptionKinds = Object.freeze([
  'json',
  'object',
  'array',
  'string',
  'number',
  'boolean',
  'null',
] satisfies readonly WorkbookActionInputDescriptionKind[])

export interface WorkbookActionInputDescription {
  readonly kind: WorkbookActionInputDescriptionKind
  readonly description?: string
  readonly required?: boolean
  readonly fields?: { readonly [key: string]: WorkbookActionInputDescription }
  readonly items?: WorkbookActionInputDescription
}

export type WorkbookActionInputIssueCode =
  | 'invalid_action_input_description'
  | 'invalid_action_input'
  | 'missing_required_input'
  | 'wrong_input_type'

export interface WorkbookActionInputIssue {
  readonly code: WorkbookActionInputIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookActionInputCheckResult =
  | {
      readonly status: 'valid'
      readonly input: WorkbookActionInput
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly input?: WorkbookActionInput
      readonly issues: readonly WorkbookActionInputIssue[]
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
  if (isWorkbookActionInputDescriptionKind(value)) {
    return value
  }
  throw new WorkbookActionInputError(`Action input description at ${path}.kind must be a supported kind`)
}

export function isWorkbookActionInputDescriptionKind(value: unknown): value is WorkbookActionInputDescriptionKind {
  return typeof value === 'string' && workbookActionInputDescriptionKinds.some((kind) => kind === value)
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

export function isWorkbookActionInputDescription(input: unknown): input is WorkbookActionInputDescription {
  try {
    normalizeWorkbookActionInputDescription(input)
    return true
  } catch {
    return false
  }
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
      return Object.freeze(output)
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
    return Object.freeze(output)
  } finally {
    seen.delete(value)
  }
}

export function normalizeWorkbookActionInput(input: unknown): WorkbookActionInput {
  return normalizeInput(input, 'input', new WeakSet())
}

export function isWorkbookActionInput(input: unknown): input is WorkbookActionInput {
  try {
    normalizeWorkbookActionInput(input)
    return true
  } catch {
    return false
  }
}

function inputIssue(code: WorkbookActionInputIssueCode, path: string, message: string): WorkbookActionInputIssue {
  return Object.freeze({ code, path, message })
}

function typeLabel(kind: WorkbookActionInputDescriptionKind): string {
  switch (kind) {
    case 'json':
      return 'JSON-safe value'
    case 'object':
      return 'an object'
    case 'array':
      return 'an array'
    case 'string':
      return 'a string'
    case 'number':
      return 'a number'
    case 'boolean':
      return 'a boolean'
    case 'null':
      return 'null'
  }
}

function matchesDescriptionKind(kind: WorkbookActionInputDescriptionKind, value: WorkbookActionInput): boolean {
  switch (kind) {
    case 'json':
      return true
    case 'object':
      return typeof value === 'object' && value !== null && !Array.isArray(value)
    case 'array':
      return Array.isArray(value)
    case 'string':
      return typeof value === 'string'
    case 'number':
      return typeof value === 'number'
    case 'boolean':
      return typeof value === 'boolean'
    case 'null':
      return value === null
  }
}

function isInputObject(value: WorkbookActionInput): value is { readonly [key: string]: WorkbookActionInput } {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateInputDescription(
  description: WorkbookActionInputDescription,
  value: WorkbookActionInput | undefined,
  path: string,
  issues: WorkbookActionInputIssue[],
): void {
  if (value === undefined) {
    if (description.required === true) {
      issues.push(inputIssue('missing_required_input', path, `Action input at ${path} is required`))
    }
    return
  }

  if (!matchesDescriptionKind(description.kind, value)) {
    issues.push(inputIssue('wrong_input_type', path, `Action input at ${path} must be ${typeLabel(description.kind)}`))
    return
  }

  if (description.kind === 'object' && description.fields !== undefined && isInputObject(value)) {
    Object.entries(description.fields).forEach(([key, fieldDescription]) => {
      const fieldPath = childPath(path, key)
      validateInputDescription(fieldDescription, Object.hasOwn(value, key) ? value[key] : undefined, fieldPath, issues)
    })
  }

  if (description.kind === 'array' && description.items !== undefined && Array.isArray(value)) {
    const itemDescription = description.items
    value.forEach((entry, index) => {
      validateInputDescription(itemDescription, entry, `${path}[${index}]`, issues)
    })
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function checkInput(description: unknown, input: unknown): WorkbookActionInputCheckResult {
  let normalizedDescription: WorkbookActionInputDescription
  try {
    normalizedDescription = normalizeWorkbookActionInputDescription(description)
  } catch (error) {
    return {
      status: 'invalid',
      issues: Object.freeze([inputIssue('invalid_action_input_description', 'input', errorMessage(error))]),
    }
  }

  let normalizedInput: WorkbookActionInput
  try {
    normalizedInput = normalizeWorkbookActionInput(input)
  } catch (error) {
    return {
      status: 'invalid',
      issues: Object.freeze([inputIssue('invalid_action_input', 'input', errorMessage(error))]),
    }
  }

  const issues: WorkbookActionInputIssue[] = []
  validateInputDescription(normalizedDescription, normalizedInput, 'input', issues)
  if (issues.length > 0) {
    return {
      status: 'invalid',
      input: normalizedInput,
      issues: Object.freeze(issues),
    }
  }

  return {
    status: 'valid',
    input: normalizedInput,
    issues: Object.freeze([]),
  }
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
