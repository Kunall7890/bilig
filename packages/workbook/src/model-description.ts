import type { WorkbookModelDescription } from './describe.js'
import { WorkbookActionInputError, normalizeWorkbookActionInputDescription, type WorkbookActionInputDescription } from './input.js'
import type { WorkbookActionInspection } from './model.js'

export type WorkbookModelDescriptionIssueCode =
  | 'invalid_type'
  | 'missing_field'
  | 'unexpected_field'
  | 'invalid_field'
  | 'duplicate_action'
  | 'action_mismatch'

export interface WorkbookModelDescriptionIssue {
  readonly code: WorkbookModelDescriptionIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookModelDescriptionCheckResult =
  | {
      readonly status: 'valid'
      readonly description: WorkbookModelDescription
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookModelDescriptionIssue[]
    }

interface ActionDescriptionBuild {
  readonly action?: WorkbookActionInspection
  readonly name?: string
}

interface ModelDescriptionBuild {
  readonly description?: WorkbookModelDescription
  readonly actionNames: readonly string[]
  readonly actionDetailNames: readonly string[]
}

function modelDescriptionIssue(code: WorkbookModelDescriptionIssueCode, path: string, message: string): WorkbookModelDescriptionIssue {
  return Object.freeze({ code, path, message })
}

function freezeModelDescriptionData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null || seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeModelDescriptionData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function dataValue(issues: WorkbookModelDescriptionIssue[], value: object, key: string, path: string, required: boolean): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    if (required) {
      issues.push(modelDescriptionIssue('missing_field', path, `Workbook model description ${path} is required`))
    }
    return undefined
  }
  if (!('value' in descriptor)) {
    issues.push(modelDescriptionIssue('invalid_field', path, `Workbook model description ${path} must be a data property`))
    return undefined
  }
  return descriptor.value
}

function pushUnexpectedFields(issues: WorkbookModelDescriptionIssue[], value: object, path: string, allowed: ReadonlySet<string>): void {
  if (Object.getOwnPropertySymbols(value).length > 0) {
    issues.push(modelDescriptionIssue('unexpected_field', path, `Workbook model description ${path} must not contain symbol keys`))
  }
  Object.keys(value).forEach((key) => {
    if (!allowed.has(key)) {
      const fieldPath = path === 'description' ? key : `${path}.${key}`
      issues.push(modelDescriptionIssue('unexpected_field', fieldPath, `Workbook model description ${fieldPath} is not supported`))
    }
  })
}

function normalizeRequiredName(issues: WorkbookModelDescriptionIssue[], value: unknown, path: string, label: string): string | undefined {
  if (typeof value !== 'string') {
    issues.push(modelDescriptionIssue('invalid_field', path, `${label} must be a string`))
    return undefined
  }
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(modelDescriptionIssue('invalid_field', path, `${label} cannot be empty`))
    return undefined
  }
  if (normalized !== value) {
    issues.push(modelDescriptionIssue('invalid_field', path, `${label} must not have leading or trailing whitespace`))
    return undefined
  }
  return normalized
}

function normalizeOptionalDescription(
  issues: WorkbookModelDescriptionIssue[],
  value: unknown,
  path: string,
  label: string,
): string | undefined {
  if (value === undefined) {
    return undefined
  }
  if (typeof value !== 'string') {
    issues.push(modelDescriptionIssue('invalid_field', path, `${label} must be a string`))
    return undefined
  }
  const normalized = value.trim()
  if (normalized === '') {
    issues.push(modelDescriptionIssue('invalid_field', path, `${label} cannot be empty`))
    return undefined
  }
  return normalized
}

function inputDescriptionPath(basePath: string, error: unknown): string {
  if (!(error instanceof WorkbookActionInputError)) {
    return basePath
  }
  if (error.path === 'input') {
    return basePath
  }
  if (error.path.startsWith('input.')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  if (error.path.startsWith('input[')) {
    return `${basePath}${error.path.slice('input'.length)}`
  }
  return basePath
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function normalizeInputDescription(
  issues: WorkbookModelDescriptionIssue[],
  value: unknown,
  path: string,
): WorkbookActionInputDescription | undefined {
  if (value === undefined) {
    return undefined
  }
  try {
    return normalizeWorkbookActionInputDescription(value)
  } catch (error) {
    issues.push(
      modelDescriptionIssue(
        'invalid_field',
        inputDescriptionPath(path, error),
        `Workbook model description ${path} input is invalid: ${errorMessage(error)}`,
      ),
    )
    return undefined
  }
}

function buildStringArray(issues: WorkbookModelDescriptionIssue[], value: unknown, path: string, label: string): readonly string[] {
  if (!Array.isArray(value)) {
    issues.push(modelDescriptionIssue('invalid_field', path, `Workbook model description ${path} must be an array`))
    return Object.freeze([])
  }
  const output: string[] = []
  const seen = new Set<string>()
  const descriptors = Object.getOwnPropertyDescriptors(value)
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${String(index)}]`
    const descriptor = descriptors[String(index)]
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(modelDescriptionIssue('invalid_field', entryPath, `Workbook model description ${entryPath} must be a data property`))
      continue
    }
    const name = normalizeRequiredName(issues, descriptor.value, entryPath, label)
    if (name === undefined) {
      continue
    }
    if (seen.has(name)) {
      issues.push(modelDescriptionIssue('duplicate_action', entryPath, `Workbook model description ${entryPath} duplicates ${name}`))
      continue
    }
    seen.add(name)
    output.push(name)
  }
  return Object.freeze(output)
}

function buildActionDescription(issues: WorkbookModelDescriptionIssue[], value: unknown, path: string): ActionDescriptionBuild {
  if (!isPlainObject(value)) {
    issues.push(modelDescriptionIssue('invalid_type', path, `Workbook model description ${path} must be a plain object`))
    return {}
  }
  pushUnexpectedFields(issues, value, path, new Set(['name', 'description', 'input']))
  const name = normalizeRequiredName(
    issues,
    dataValue(issues, value, 'name', `${path}.name`, true),
    `${path}.name`,
    `Workbook model description ${path}.name`,
  )
  const description = normalizeOptionalDescription(
    issues,
    dataValue(issues, value, 'description', `${path}.description`, false),
    `${path}.description`,
    `Workbook model description ${path}.description`,
  )
  const input = normalizeInputDescription(issues, dataValue(issues, value, 'input', `${path}.input`, false), `${path}.input`)
  if (name === undefined) {
    return {}
  }
  return {
    name,
    action: freezeModelDescriptionData({
      name,
      ...(description !== undefined ? { description } : {}),
      ...(input !== undefined ? { input } : {}),
    }),
  }
}

function buildActionDescriptions(
  issues: WorkbookModelDescriptionIssue[],
  value: unknown,
  path: string,
): { readonly actions: readonly WorkbookActionInspection[]; readonly names: readonly string[] } {
  if (!Array.isArray(value)) {
    issues.push(modelDescriptionIssue('invalid_field', path, `Workbook model description ${path} must be an array`))
    return { actions: Object.freeze([]), names: Object.freeze([]) }
  }
  const actions: WorkbookActionInspection[] = []
  const names: string[] = []
  const seen = new Set<string>()
  const descriptors = Object.getOwnPropertyDescriptors(value)
  for (let index = 0; index < value.length; index += 1) {
    const entryPath = `${path}[${String(index)}]`
    const descriptor = descriptors[String(index)]
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(modelDescriptionIssue('invalid_field', entryPath, `Workbook model description ${entryPath} must be a data property`))
      continue
    }
    const built = buildActionDescription(issues, descriptor.value, entryPath)
    if (built.name === undefined || built.action === undefined) {
      continue
    }
    if (seen.has(built.name)) {
      issues.push(
        modelDescriptionIssue(
          'duplicate_action',
          `${entryPath}.name`,
          `Workbook model description ${entryPath}.name duplicates ${built.name}`,
        ),
      )
      continue
    }
    seen.add(built.name)
    names.push(built.name)
    actions.push(built.action)
  }
  return {
    actions: freezeModelDescriptionData(actions),
    names: Object.freeze(names),
  }
}

function actionNamesMatch(left: readonly string[], right: readonly string[]): boolean {
  return left.length === right.length && left.every((name, index) => right[index] === name)
}

function buildModelDescription(value: unknown, issues: WorkbookModelDescriptionIssue[]): ModelDescriptionBuild {
  if (!isPlainObject(value)) {
    issues.push(modelDescriptionIssue('invalid_type', 'description', 'Workbook model description must be a plain object'))
    return {
      actionNames: Object.freeze([]),
      actionDetailNames: Object.freeze([]),
    }
  }

  pushUnexpectedFields(issues, value, 'description', new Set(['name', 'description', 'actions', 'actionDetails', 'hasChecks']))
  const name = normalizeRequiredName(issues, dataValue(issues, value, 'name', 'name', true), 'name', 'Workbook model description name')
  const description = normalizeOptionalDescription(
    issues,
    dataValue(issues, value, 'description', 'description', false),
    'description',
    'Workbook model description description',
  )
  const actionNames = buildStringArray(
    issues,
    dataValue(issues, value, 'actions', 'actions', true),
    'actions',
    'Workbook model description action name',
  )
  const actionDetails = buildActionDescriptions(issues, dataValue(issues, value, 'actionDetails', 'actionDetails', true), 'actionDetails')
  const hasChecksValue = dataValue(issues, value, 'hasChecks', 'hasChecks', true)
  const hasChecks = typeof hasChecksValue === 'boolean' ? hasChecksValue : undefined
  if (hasChecksValue !== undefined && typeof hasChecksValue !== 'boolean') {
    issues.push(modelDescriptionIssue('invalid_field', 'hasChecks', 'Workbook model description hasChecks must be a boolean'))
  }
  if (!actionNamesMatch(actionNames, actionDetails.names)) {
    issues.push(
      modelDescriptionIssue(
        'action_mismatch',
        'actionDetails',
        'Workbook model description actionDetails names must exactly match actions',
      ),
    )
  }
  if (name === undefined || hasChecks === undefined) {
    return {
      actionNames,
      actionDetailNames: actionDetails.names,
    }
  }
  return {
    actionNames,
    actionDetailNames: actionDetails.names,
    description: freezeModelDescriptionData({
      name,
      ...(description !== undefined ? { description } : {}),
      actions: actionNames,
      actionDetails: actionDetails.actions,
      hasChecks,
    }),
  }
}

export function checkWorkbookModelDescription(value: unknown): WorkbookModelDescriptionCheckResult {
  const issues: WorkbookModelDescriptionIssue[] = []
  const built = buildModelDescription(value, issues)
  if (issues.length > 0 || built.description === undefined) {
    return freezeModelDescriptionData({
      status: 'invalid',
      issues,
    })
  }
  return freezeModelDescriptionData({
    status: 'valid',
    description: built.description,
    issues: [],
  })
}

export function isWorkbookModelDescription(value: unknown): value is WorkbookModelDescription {
  return checkWorkbookModelDescription(value).status === 'valid'
}
