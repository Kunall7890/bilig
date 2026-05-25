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
  readonly values?: readonly WorkbookActionInput[]
  readonly min?: number
  readonly max?: number
  readonly minLength?: number
  readonly maxLength?: number
  readonly pattern?: string
  readonly minItems?: number
  readonly maxItems?: number
  readonly additionalProperties?: boolean
  readonly default?: WorkbookActionInput
  readonly examples?: readonly WorkbookActionInput[]
}

export type WorkbookActionInputIssueCode =
  | 'invalid_action_input_description'
  | 'invalid_action_input'
  | 'missing_required_input'
  | 'wrong_input_type'
  | 'unknown_input_field'
  | 'input_constraint_failed'

export interface WorkbookActionInputIssue {
  readonly code: WorkbookActionInputIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookActionInputCheckResult =
  | {
      readonly status: 'valid'
      readonly input?: WorkbookActionInput
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly input?: WorkbookActionInput
      readonly issues: readonly WorkbookActionInputIssue[]
    }

export class WorkbookActionInputError extends Error {
  readonly path: string

  constructor(message: string, path = 'input') {
    super(message)
    this.name = 'WorkbookActionInputError'
    this.path = path
  }
}

function inputError(path: string, message: string): WorkbookActionInputError {
  return new WorkbookActionInputError(message, path)
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

function defineOwnDataProperty(target: object, key: string, value: unknown): void {
  Object.defineProperty(target, key, {
    value,
    enumerable: true,
    configurable: true,
    writable: true,
  })
}

function ownDataValue(value: object, key: string, path: string, label: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return undefined
  }
  if (!('value' in descriptor)) {
    throw inputError(path, `${label} at ${path} must be a data property`)
  }
  return descriptor.value
}

function sortedOwnEnumerableDataEntries(value: object, path: string, label: string): readonly (readonly [string, unknown])[] {
  const entries: [string, unknown][] = []
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable) {
      return
    }
    const entryPath = childPath(path, key)
    if (!('value' in descriptor)) {
      throw inputError(entryPath, `${label} at ${entryPath} must be a data property`)
    }
    entries.push([key, descriptor.value])
  })
  return entries.toSorted(([left], [right]) => left.localeCompare(right))
}

function inputDescriptionKind(value: unknown, path: string): WorkbookActionInputDescriptionKind {
  if (isWorkbookActionInputDescriptionKind(value)) {
    return value
  }
  throw inputError(`${path}.kind`, `Action input description at ${path}.kind must be a supported kind`)
}

export function isWorkbookActionInputDescriptionKind(value: unknown): value is WorkbookActionInputDescriptionKind {
  return typeof value === 'string' && workbookActionInputDescriptionKinds.some((kind) => kind === value)
}

function normalizeDescriptionText(value: unknown, path: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'string') {
    throw inputError(path, `Action input description at ${path} must be a string`)
  }
  const text = value.trim()
  if (text === '') {
    throw inputError(path, `Action input description at ${path} cannot be empty`)
  }
  return text
}

function normalizeDescriptionBoolean(value: unknown, path: string): boolean | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'boolean') {
    throw inputError(path, `Action input description at ${path} must be a boolean`)
  }
  return value
}

function normalizeFiniteDescriptionNumber(value: unknown, path: string): number | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw inputError(path, `Action input description at ${path} must be a finite number`)
  }
  return value
}

function normalizeNonNegativeInteger(value: unknown, path: string): number | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) {
    throw inputError(path, `Action input description at ${path} must be a non-negative safe integer`)
  }
  return value
}

function normalizePattern(value: unknown, path: string): string | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'string') {
    throw inputError(path, `Action input description at ${path} must be a string`)
  }
  try {
    RegExp(value)
  } catch (error) {
    throw inputError(path, `Action input description at ${path} must be a valid regular expression: ${errorMessage(error)}`)
  }
  return value
}

function normalizeDescriptionInputArray(value: unknown, path: string): readonly WorkbookActionInput[] | undefined {
  if (value === undefined) {
    return undefined
  }
  if (!Array.isArray(value)) {
    throw inputError(path, `Action input description at ${path} must be an array`)
  }
  if (value.length === 0) {
    throw inputError(path, `Action input description at ${path} cannot be empty`)
  }
  const output: WorkbookActionInput[] = []
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${String(index)}]`
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable) {
      throw inputError(entryPath, `Action input description at ${entryPath} must be a data property`)
    }
    if (!('value' in descriptor)) {
      throw inputError(entryPath, `Action input description at ${entryPath} must be a data property`)
    }
    output.push(normalizeInput(descriptor.value, entryPath, new WeakSet()))
  }
  return Object.freeze(output)
}

function assertUniqueAllowedValues(values: readonly WorkbookActionInput[] | undefined, path: string): void {
  if (values === undefined) {
    return
  }
  const seen = new Map<string, number>()
  values.forEach((value, index) => {
    const key = canonicalInputJson(value)
    const firstIndex = seen.get(key)
    if (firstIndex !== undefined) {
      throw inputError(
        `${path}.values[${String(index)}]`,
        `Action input description at ${path}.values[${String(index)}] duplicates ${path}.values[${String(firstIndex)}]`,
      )
    }
    seen.set(key, index)
  })
}

function normalizeInputDescription(value: unknown, path: string, seen: WeakSet<object>): WorkbookActionInputDescription {
  if (!isPlainObject(value)) {
    throw inputError(path, `Action input description at ${path} must be a plain object`)
  }

  if (seen.has(value)) {
    throw inputError(path, `Action input description at ${path} must not contain cycles`)
  }
  seen.add(value)

  try {
    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw inputError(path, `Action input description at ${path} must not contain symbol keys`)
    }

    const allowedKeys = new Set([
      'kind',
      'description',
      'required',
      'fields',
      'items',
      'values',
      'min',
      'max',
      'minLength',
      'maxLength',
      'pattern',
      'minItems',
      'maxItems',
      'additionalProperties',
      'default',
      'examples',
    ])
    const unknownKey = Object.keys(value).find((key) => !allowedKeys.has(key))
    if (unknownKey !== undefined) {
      const unknownPath = childPath(path, unknownKey)
      throw inputError(unknownPath, `Action input description at ${unknownPath} is not supported`)
    }

    const kind = inputDescriptionKind(ownDataValue(value, 'kind', `${path}.kind`, 'Action input description'), path)
    const output: {
      kind: WorkbookActionInputDescriptionKind
      description?: string
      required?: boolean
      fields?: { readonly [key: string]: WorkbookActionInputDescription }
      items?: WorkbookActionInputDescription
      values?: readonly WorkbookActionInput[]
      min?: number
      max?: number
      minLength?: number
      maxLength?: number
      pattern?: string
      minItems?: number
      maxItems?: number
      additionalProperties?: boolean
      default?: WorkbookActionInput
      examples?: readonly WorkbookActionInput[]
    } = { kind }

    const description = normalizeDescriptionText(
      ownDataValue(value, 'description', `${path}.description`, 'Action input description'),
      `${path}.description`,
    )
    if (description !== undefined) {
      output.description = description
    }

    const required = normalizeDescriptionBoolean(
      ownDataValue(value, 'required', `${path}.required`, 'Action input description'),
      `${path}.required`,
    )
    if (required !== undefined) {
      output.required = required
    }

    const fieldsValue = ownDataValue(value, 'fields', `${path}.fields`, 'Action input description')
    if (fieldsValue !== undefined) {
      if (kind !== 'object') {
        throw inputError(`${path}.fields`, `Action input description at ${path}.fields can only be used when kind is object`)
      }
      if (!isPlainObject(fieldsValue)) {
        throw inputError(`${path}.fields`, `Action input description at ${path}.fields must be a plain object`)
      }
      const fields: Record<string, WorkbookActionInputDescription> = {}
      sortedOwnEnumerableDataEntries(fieldsValue, `${path}.fields`, 'Action input description').forEach(([key, entry]) => {
        if (key.trim() === '') {
          throw inputError(`${path}.fields`, `Action input description at ${path}.fields cannot contain an empty field name`)
        }
        defineOwnDataProperty(fields, key, normalizeInputDescription(entry, childPath(`${path}.fields`, key), seen))
      })
      output.fields = Object.freeze(fields)
    }

    const items = ownDataValue(value, 'items', `${path}.items`, 'Action input description')
    if (items !== undefined) {
      if (kind !== 'array') {
        throw inputError(`${path}.items`, `Action input description at ${path}.items can only be used when kind is array`)
      }
      output.items = normalizeInputDescription(items, `${path}.items`, seen)
    }

    const min = normalizeFiniteDescriptionNumber(ownDataValue(value, 'min', `${path}.min`, 'Action input description'), `${path}.min`)
    const max = normalizeFiniteDescriptionNumber(ownDataValue(value, 'max', `${path}.max`, 'Action input description'), `${path}.max`)
    if (min !== undefined || max !== undefined) {
      if (kind !== 'number') {
        throw inputError(`${path}.min`, `Action input description number bounds can only be used when kind is number`)
      }
      if (min !== undefined) {
        output.min = min
      }
      if (max !== undefined) {
        output.max = max
      }
      if (min !== undefined && max !== undefined && min > max) {
        throw inputError(`${path}.min`, `Action input description at ${path}.min must be less than or equal to ${path}.max`)
      }
    }

    const minLength = normalizeNonNegativeInteger(
      ownDataValue(value, 'minLength', `${path}.minLength`, 'Action input description'),
      `${path}.minLength`,
    )
    const maxLength = normalizeNonNegativeInteger(
      ownDataValue(value, 'maxLength', `${path}.maxLength`, 'Action input description'),
      `${path}.maxLength`,
    )
    const pattern = normalizePattern(ownDataValue(value, 'pattern', `${path}.pattern`, 'Action input description'), `${path}.pattern`)
    if (minLength !== undefined || maxLength !== undefined || pattern !== undefined) {
      if (kind !== 'string') {
        throw inputError(`${path}.minLength`, `Action input description string constraints can only be used when kind is string`)
      }
      if (minLength !== undefined) {
        output.minLength = minLength
      }
      if (maxLength !== undefined) {
        output.maxLength = maxLength
      }
      if (pattern !== undefined) {
        output.pattern = pattern
      }
      if (minLength !== undefined && maxLength !== undefined && minLength > maxLength) {
        throw inputError(
          `${path}.minLength`,
          `Action input description at ${path}.minLength must be less than or equal to ${path}.maxLength`,
        )
      }
    }

    const minItems = normalizeNonNegativeInteger(
      ownDataValue(value, 'minItems', `${path}.minItems`, 'Action input description'),
      `${path}.minItems`,
    )
    const maxItems = normalizeNonNegativeInteger(
      ownDataValue(value, 'maxItems', `${path}.maxItems`, 'Action input description'),
      `${path}.maxItems`,
    )
    if (minItems !== undefined || maxItems !== undefined) {
      if (kind !== 'array') {
        throw inputError(`${path}.minItems`, `Action input description array constraints can only be used when kind is array`)
      }
      if (minItems !== undefined) {
        output.minItems = minItems
      }
      if (maxItems !== undefined) {
        output.maxItems = maxItems
      }
      if (minItems !== undefined && maxItems !== undefined && minItems > maxItems) {
        throw inputError(`${path}.minItems`, `Action input description at ${path}.minItems must be less than or equal to ${path}.maxItems`)
      }
    }

    const additionalProperties = normalizeDescriptionBoolean(
      ownDataValue(value, 'additionalProperties', `${path}.additionalProperties`, 'Action input description'),
      `${path}.additionalProperties`,
    )
    if (additionalProperties !== undefined) {
      if (kind !== 'object') {
        throw inputError(
          `${path}.additionalProperties`,
          `Action input description at ${path}.additionalProperties can only be used when kind is object`,
        )
      }
      output.additionalProperties = additionalProperties
    }

    const values = normalizeDescriptionInputArray(
      ownDataValue(value, 'values', `${path}.values`, 'Action input description'),
      `${path}.values`,
    )
    if (values !== undefined) {
      output.values = values
    }
    assertUniqueAllowedValues(values, path)

    const defaultValue = ownDataValue(value, 'default', `${path}.default`, 'Action input description')
    if (defaultValue !== undefined) {
      output.default = normalizeInput(defaultValue, `${path}.default`, new WeakSet())
    }

    const examples = normalizeDescriptionInputArray(
      ownDataValue(value, 'examples', `${path}.examples`, 'Action input description'),
      `${path}.examples`,
    )
    if (examples !== undefined) {
      output.examples = examples
    }

    assertDescriptionPayloads(output, path)

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
      throw inputError(path, `Action input at ${path} must be a finite number`)
    }
    return value
  }

  if (value === undefined) {
    throw inputError(path, `Action input at ${path} must not be undefined`)
  }

  if (typeof value === 'bigint' || typeof value === 'function' || typeof value === 'symbol') {
    throw inputError(path, `Action input at ${path} must be JSON-safe, not ${typeof value}`)
  }

  if (typeof value !== 'object') {
    throw inputError(path, `Action input at ${path} must be JSON-safe`)
  }

  if (seen.has(value)) {
    throw inputError(path, `Action input at ${path} must not contain cycles`)
  }
  seen.add(value)

  try {
    if (Array.isArray(value)) {
      const output: WorkbookActionInput[] = []
      for (let index = 0; index < value.length; index += 1) {
        const itemPath = `${path}[${index}]`
        const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
        if (descriptor === undefined) {
          throw inputError(`${path}[${index}]`, `Action input at ${path}[${index}] must not be a sparse array hole`)
        }
        if (!('value' in descriptor)) {
          throw inputError(itemPath, `Action input at ${itemPath} must be a data property`)
        }
        output.push(normalizeInput(descriptor.value, itemPath, seen))
      }
      return Object.freeze(output)
    }

    if (!isPlainObject(value)) {
      throw inputError(path, `Action input at ${path} must be a plain JSON object, not ${typeName(value)}`)
    }

    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw inputError(path, `Action input at ${path} must not contain symbol keys`)
    }

    const output: { [key: string]: WorkbookActionInput } = {}
    sortedOwnEnumerableDataEntries(value, path, 'Action input').forEach(([key, entry]) => {
      defineOwnDataProperty(output, key, normalizeInput(entry, childPath(path, key), seen))
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

function canonicalInputJson(value: WorkbookActionInput): string {
  return JSON.stringify(value)
}

function pushConstraintIssue(path: string, message: string, issues: WorkbookActionInputIssue[]): void {
  issues.push(inputIssue('input_constraint_failed', path, message))
}

function validateInputConstraints(
  description: WorkbookActionInputDescription,
  value: WorkbookActionInput,
  path: string,
  issues: WorkbookActionInputIssue[],
): void {
  if (description.values !== undefined) {
    const allowed = new Set(description.values.map(canonicalInputJson))
    if (!allowed.has(canonicalInputJson(value))) {
      pushConstraintIssue(path, `Action input at ${path} must be one of the allowed values`, issues)
    }
  }
  if (description.kind === 'number' && typeof value === 'number') {
    if (description.min !== undefined && value < description.min) {
      pushConstraintIssue(path, `Action input at ${path} must be greater than or equal to ${String(description.min)}`, issues)
    }
    if (description.max !== undefined && value > description.max) {
      pushConstraintIssue(path, `Action input at ${path} must be less than or equal to ${String(description.max)}`, issues)
    }
  }
  if (description.kind === 'string' && typeof value === 'string') {
    if (description.minLength !== undefined && value.length < description.minLength) {
      pushConstraintIssue(path, `Action input at ${path} must be at least ${String(description.minLength)} characters`, issues)
    }
    if (description.maxLength !== undefined && value.length > description.maxLength) {
      pushConstraintIssue(path, `Action input at ${path} must be at most ${String(description.maxLength)} characters`, issues)
    }
    if (description.pattern !== undefined && !new RegExp(description.pattern).test(value)) {
      pushConstraintIssue(path, `Action input at ${path} must match pattern ${description.pattern}`, issues)
    }
  }
  if (description.kind === 'array' && Array.isArray(value)) {
    if (description.minItems !== undefined && value.length < description.minItems) {
      pushConstraintIssue(path, `Action input at ${path} must contain at least ${String(description.minItems)} items`, issues)
    }
    if (description.maxItems !== undefined && value.length > description.maxItems) {
      pushConstraintIssue(path, `Action input at ${path} must contain at most ${String(description.maxItems)} items`, issues)
    }
  }
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

  validateInputConstraints(description, value, path, issues)

  if (description.kind === 'object' && description.fields !== undefined && isInputObject(value)) {
    Object.entries(description.fields).forEach(([key, fieldDescription]) => {
      const fieldPath = childPath(path, key)
      validateInputDescription(fieldDescription, Object.hasOwn(value, key) ? value[key] : undefined, fieldPath, issues)
    })
  }

  if (description.kind === 'object' && description.additionalProperties === false && isInputObject(value)) {
    const fieldNames = new Set(Object.keys(description.fields ?? {}))
    Object.keys(value).forEach((key) => {
      if (!fieldNames.has(key)) {
        const fieldPath = childPath(path, key)
        issues.push(inputIssue('unknown_input_field', fieldPath, `Action input at ${fieldPath} is not allowed`))
      }
    })
  }

  if (description.kind === 'array' && description.items !== undefined && Array.isArray(value)) {
    const itemDescription = description.items
    value.forEach((entry, index) => {
      validateInputDescription(itemDescription, entry, `${path}[${index}]`, issues)
    })
  }
}

function assertDescriptionValue(description: WorkbookActionInputDescription, value: WorkbookActionInput, path: string): void {
  const issues: WorkbookActionInputIssue[] = []
  validateInputDescription(description, value, path, issues)
  const [firstIssue] = issues
  if (firstIssue !== undefined) {
    throw inputError(path, `Action input description at ${path} is invalid: ${firstIssue.message}`)
  }
}

function assertDescriptionPayloads(description: WorkbookActionInputDescription, path: string): void {
  description.values?.forEach((value, index) => {
    assertDescriptionValue(description, value, `${path}.values[${String(index)}]`)
  })
  if (description.default !== undefined) {
    assertDescriptionValue(description, description.default, `${path}.default`)
  }
  description.examples?.forEach((value, index) => {
    assertDescriptionValue(description, value, `${path}.examples[${String(index)}]`)
  })
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function errorPath(error: unknown): string {
  return error instanceof WorkbookActionInputError ? error.path : 'input'
}

export function checkInput(description: unknown, input: unknown): WorkbookActionInputCheckResult {
  let normalizedDescription: WorkbookActionInputDescription
  try {
    normalizedDescription = normalizeWorkbookActionInputDescription(description)
  } catch (error) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([inputIssue('invalid_action_input_description', errorPath(error), errorMessage(error))]),
    })
  }

  if (input === undefined) {
    const issues: WorkbookActionInputIssue[] = []
    validateInputDescription(normalizedDescription, undefined, 'input', issues)
    if (issues.length > 0) {
      return Object.freeze({
        status: 'invalid',
        issues: Object.freeze(issues),
      })
    }
    return Object.freeze({
      status: 'valid',
      issues: Object.freeze([] as const),
    })
  }

  let normalizedInput: WorkbookActionInput
  try {
    normalizedInput = normalizeWorkbookActionInput(input)
  } catch (error) {
    return Object.freeze({
      status: 'invalid',
      issues: Object.freeze([inputIssue('invalid_action_input', errorPath(error), errorMessage(error))]),
    })
  }

  const issues: WorkbookActionInputIssue[] = []
  validateInputDescription(normalizedDescription, normalizedInput, 'input', issues)
  if (issues.length > 0) {
    return Object.freeze({
      status: 'invalid',
      input: normalizedInput,
      issues: Object.freeze(issues),
    })
  }

  return Object.freeze({
    status: 'valid',
    input: normalizedInput,
    issues: Object.freeze([] as const),
  })
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
