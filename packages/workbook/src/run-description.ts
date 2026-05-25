import type { WorkbookRunResultDescription } from './describe.js'
import { isWorkbookRunErrorCode } from './result.js'

export type WorkbookRunResultDescriptionIssueCode =
  | 'invalid_type'
  | 'invalid_status'
  | 'missing_field'
  | 'unexpected_field'
  | 'invalid_field'

export interface WorkbookRunResultDescriptionIssue {
  readonly code: WorkbookRunResultDescriptionIssueCode
  readonly path: string
  readonly message: string
}

export type WorkbookRunResultDescriptionCheckResult =
  | {
      readonly status: 'valid'
      readonly description: WorkbookRunResultDescription
      readonly issues: readonly []
    }
  | {
      readonly status: 'invalid'
      readonly issues: readonly WorkbookRunResultDescriptionIssue[]
    }

function runResultDescriptionIssue(
  code: WorkbookRunResultDescriptionIssueCode,
  path: string,
  message: string,
): WorkbookRunResultDescriptionIssue {
  return {
    code,
    path,
    message,
  }
}

function freezeResultDescriptionData<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeResultDescriptionData(descriptor.value, seen)
    }
  })
  Object.freeze(value)
  return value
}

function cloneResultDescriptionData<T>(value: T, seen?: WeakMap<object, unknown>): T
function cloneResultDescriptionData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
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
      if (descriptor !== undefined && descriptor.enumerable && 'value' in descriptor) {
        cloned.push(cloneResultDescriptionData(descriptor.value, seen))
      }
    }
    return cloned
  }

  const cloned: Record<string, unknown> = {}
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable || !('value' in descriptor)) {
      return
    }
    Object.defineProperty(cloned, key, {
      configurable: true,
      enumerable: true,
      value: cloneResultDescriptionData(descriptor.value, seen),
      writable: true,
    })
  })
  return cloned
}

function ownJsonDescriptionValue(value: Record<string, unknown>, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor.value : undefined
}

function hasOptionalJsonDescriptionValue(value: Record<string, unknown>, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return true
  }
  return 'value' in descriptor && isJsonDescriptionValue(descriptor.value)
}

function hasJsonDescriptionArray(value: Record<string, unknown>, key: string): boolean {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor && Array.isArray(descriptor.value) && isJsonDescriptionValue(descriptor.value)
}

function isJsonDescriptionObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function pushJsonDescriptionValueIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  value: unknown,
  path: string,
  seen = new WeakSet<object>(),
): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      issues.push(runResultDescriptionIssue('invalid_field', path, `Workbook run result description ${path} must be JSON-safe`))
    }
    return
  }
  if (typeof value !== 'object') {
    issues.push(runResultDescriptionIssue('invalid_field', path, `Workbook run result description ${path} must be JSON-safe`))
    return
  }
  if (seen.has(value)) {
    issues.push(runResultDescriptionIssue('invalid_field', path, `Workbook run result description ${path} must not be circular`))
    return
  }
  seen.add(value)

  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const entryPath = `${path}[${index}]`
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        issues.push(
          runResultDescriptionIssue('invalid_field', entryPath, `Workbook run result description ${entryPath} must be a data property`),
        )
        continue
      }
      pushJsonDescriptionValueIssues(issues, descriptor.value, entryPath, seen)
    }
    return
  }

  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', path, `Workbook run result description ${path} must be a plain object`))
    return
  }
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (!descriptor.enumerable) {
      continue
    }
    const entryPath = `${path}.${key}`
    if (!('value' in descriptor)) {
      issues.push(
        runResultDescriptionIssue('invalid_field', entryPath, `Workbook run result description ${entryPath} must be a data property`),
      )
      continue
    }
    pushJsonDescriptionValueIssues(issues, descriptor.value, entryPath, seen)
  }
}

function isJsonDescriptionValue(value: unknown, seen = new WeakSet<object>()): boolean {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return true
  }
  if (typeof value === 'number') {
    return Number.isFinite(value)
  }
  if (typeof value !== 'object') {
    return false
  }
  if (seen.has(value)) {
    return false
  }
  seen.add(value)

  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        return false
      }
      if (!isJsonDescriptionValue(descriptor.value, seen)) {
        return false
      }
    }
    return true
  }

  if (!isJsonDescriptionObject(value)) {
    return false
  }
  for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(value))) {
    if (!descriptor.enumerable) {
      continue
    }
    if (!('value' in descriptor) || !isJsonDescriptionValue(descriptor.value, seen)) {
      return false
    }
  }
  return true
}

function pushUnexpectedRunDescriptionFields(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  allowedFields: ReadonlySet<string>,
): void {
  for (const key of Object.keys(value)) {
    if (!allowedFields.has(key)) {
      issues.push(runResultDescriptionIssue('unexpected_field', key, `Workbook run result description must not include ${key}`))
    }
  }
}

function pushUnexpectedObjectFields(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  allowedFields: ReadonlySet<string>,
): void {
  for (const key of Object.keys(value)) {
    if (!allowedFields.has(key)) {
      issues.push(
        runResultDescriptionIssue('unexpected_field', `${path}.${key}`, `Workbook run result description ${path} must not include ${key}`),
      )
    }
  }
}

function ownRunDescriptionDataValue(value: Record<string, unknown>, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor.value : undefined
}

function pushRequiredStringFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(runResultDescriptionIssue('missing_field', fieldPath, `Workbook run result description ${fieldPath} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a string`))
  }
}

function pushOptionalStringFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a string`))
  }
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function pushSafeNonNegativeIntegerFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (!isSafeNonNegativeInteger(descriptor.value)) {
    issues.push(
      runResultDescriptionIssue(
        'invalid_field',
        fieldPath,
        `Workbook run result description ${fieldPath} must be a non-negative safe integer`,
      ),
    )
  }
}

function pushRequiredSafeNonNegativeIntegerFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(runResultDescriptionIssue('missing_field', fieldPath, `Workbook run result description ${fieldPath} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (!isSafeNonNegativeInteger(descriptor.value)) {
    issues.push(
      runResultDescriptionIssue(
        'invalid_field',
        fieldPath,
        `Workbook run result description ${fieldPath} must be a non-negative safe integer`,
      ),
    )
  }
}

function pushRequiredJsonArrayFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(runResultDescriptionIssue('missing_field', fieldPath, `Workbook run result description ${fieldPath} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (!Array.isArray(descriptor.value)) {
    issues.push(runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be an array`))
  }
}

function pushJsonArrayFieldIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(
      runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`),
    )
    return
  }
  if (!Array.isArray(descriptor.value)) {
    issues.push(runResultDescriptionIssue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be an array`))
  }
}

function pushRunDescriptionEntryIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  value: unknown,
  path: string,
  validator: (entry: Record<string, unknown>, entryPath: string) => void,
): void {
  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', path, `Workbook run result description ${path} must be a plain object`))
    return
  }
  validator(value, path)
}

function pushRunDescriptionArrayEntryIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  key: string,
  validator: (entry: Record<string, unknown>, entryPath: string) => void,
): void {
  const array = ownRunDescriptionDataValue(value, key)
  if (!Array.isArray(array)) {
    return
  }
  const descriptors = Object.getOwnPropertyDescriptors(array)
  for (let index = 0; index < array.length; index += 1) {
    const descriptor = descriptors[String(index)]
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      continue
    }
    pushRunDescriptionEntryIssues(issues, descriptor.value, `${key}[${index}]`, validator)
  }
}

function pushChangeDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], entry: Record<string, unknown>, path: string): void {
  pushUnexpectedObjectFields(issues, entry, path, new Set(['kind', 'target', 'message']))
  pushRequiredStringFieldIssue(issues, entry, path, 'kind')
  pushRequiredStringFieldIssue(issues, entry, path, 'message')
}

function pushCheckDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], entry: Record<string, unknown>, path: string): void {
  pushUnexpectedObjectFields(issues, entry, path, new Set(['status', 'kind', 'target', 'refs', 'message', 'expectation', 'proof']))
  const status = ownRunDescriptionDataValue(entry, 'status')
  if (status !== 'planned' && status !== 'passed' && status !== 'failed') {
    issues.push(
      runResultDescriptionIssue(
        'invalid_field',
        `${path}.status`,
        `Workbook run result description ${path}.status must be planned, passed, or failed`,
      ),
    )
  }
  pushRequiredStringFieldIssue(issues, entry, path, 'kind')
  pushRequiredStringFieldIssue(issues, entry, path, 'message')
}

function pushErrorDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], entry: Record<string, unknown>, path: string): void {
  pushUnexpectedObjectFields(issues, entry, path, new Set(['code', 'message', 'path', 'issueCode']))
  const code = ownRunDescriptionDataValue(entry, 'code')
  if (!isWorkbookRunErrorCode(code)) {
    issues.push(
      runResultDescriptionIssue(
        'invalid_field',
        `${path}.code`,
        `Workbook run result description ${path}.code must be a workbook run error code`,
      ),
    )
  }
  pushRequiredStringFieldIssue(issues, entry, path, 'message')
  pushOptionalStringFieldIssue(issues, entry, path, 'path')
  pushOptionalStringFieldIssue(issues, entry, path, 'issueCode')
}

function pushUnverifiedDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], entry: Record<string, unknown>, path: string): void {
  pushUnexpectedObjectFields(issues, entry, path, new Set(['kind', 'message']))
  const kind = ownRunDescriptionDataValue(entry, 'kind')
  if (kind !== 'apply' && kind !== 'plan') {
    issues.push(
      runResultDescriptionIssue('invalid_field', `${path}.kind`, `Workbook run result description ${path}.kind must be apply or plan`),
    )
  }
  pushRequiredStringFieldIssue(issues, entry, path, 'message')
}

function pushCommandReceiptDescriptionIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  entry: Record<string, unknown>,
  path: string,
): void {
  pushUnexpectedObjectFields(
    issues,
    entry,
    path,
    new Set(['commandIndex', 'commandKind', 'commandDigest', 'previewOps', 'appliedOps', 'resolvedRefs', 'formulaLabels', 'proof']),
  )
  pushRequiredSafeNonNegativeIntegerFieldIssue(issues, entry, path, 'commandIndex')
  pushRequiredStringFieldIssue(issues, entry, path, 'commandKind')
  pushRequiredStringFieldIssue(issues, entry, path, 'commandDigest')
  pushRequiredJsonArrayFieldIssue(issues, entry, path, 'previewOps')
  pushRequiredJsonArrayFieldIssue(issues, entry, path, 'appliedOps')
  const resolvedRefs = ownJsonDescriptionValue(entry, 'resolvedRefs')
  if (resolvedRefs !== undefined) {
    pushCommandResolvedRefsIssues(issues, resolvedRefs, `${path}.resolvedRefs`)
  }
  pushJsonArrayFieldIssue(issues, entry, path, 'formulaLabels')
}

function pushConcreteRefDataIssues(issues: WorkbookRunResultDescriptionIssue[], value: unknown, path: string): void {
  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', path, `Workbook run result description ${path} must be concrete range ref data`))
    return
  }
  pushUnexpectedObjectFields(issues, value, path, new Set(['kind', 'id', 'label', 'range']))
  const kind = ownJsonDescriptionValue(value, 'kind')
  if (kind !== 'range') {
    issues.push(runResultDescriptionIssue('invalid_field', `${path}.kind`, `Workbook run result description ${path}.kind must be range`))
  }
  pushRequiredStringFieldIssue(issues, value, path, 'id')
  pushRequiredStringFieldIssue(issues, value, path, 'label')
  const range = ownJsonDescriptionValue(value, 'range')
  if (!isJsonDescriptionObject(range)) {
    issues.push(runResultDescriptionIssue('invalid_type', `${path}.range`, `Workbook run result description ${path}.range must be a range`))
    return
  }
  pushUnexpectedObjectFields(issues, range, `${path}.range`, new Set(['sheetName', 'startAddress', 'endAddress']))
  pushRequiredStringFieldIssue(issues, range, `${path}.range`, 'sheetName')
  pushRequiredStringFieldIssue(issues, range, `${path}.range`, 'startAddress')
  pushRequiredStringFieldIssue(issues, range, `${path}.range`, 'endAddress')
}

function pushResolvedRefValueIssues(issues: WorkbookRunResultDescriptionIssue[], value: unknown, path: string): void {
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        issues.push(
          runResultDescriptionIssue(
            'invalid_field',
            `${path}[${String(index)}]`,
            `Workbook run result description ${path}[${String(index)}] must be a data property`,
          ),
        )
        continue
      }
      pushConcreteRefDataIssues(issues, descriptor.value, `${path}[${String(index)}]`)
    }
    return
  }
  pushConcreteRefDataIssues(issues, value, path)
}

function pushCommandResolvedRefsIssues(issues: WorkbookRunResultDescriptionIssue[], value: unknown, path: string): void {
  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', path, `Workbook run result description ${path} must be an object`))
    return
  }
  pushUnexpectedObjectFields(issues, value, path, new Set(['target', 'inputs']))
  const target = ownJsonDescriptionValue(value, 'target')
  if (target !== undefined) {
    pushResolvedRefValueIssues(issues, target, `${path}.target`)
  }
  const inputs = ownJsonDescriptionValue(value, 'inputs')
  if (inputs === undefined) {
    return
  }
  if (!Array.isArray(inputs)) {
    issues.push(
      runResultDescriptionIssue('invalid_type', `${path}.inputs`, `Workbook run result description ${path}.inputs must be an array`),
    )
    return
  }
  for (let index = 0; index < inputs.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(inputs, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
      issues.push(
        runResultDescriptionIssue(
          'invalid_field',
          `${path}.inputs[${String(index)}]`,
          `Workbook run result description ${path}.inputs[${String(index)}] must be a data property`,
        ),
      )
      continue
    }
    pushResolvedRefValueIssues(issues, descriptor.value, `${path}.inputs[${String(index)}]`)
  }
}

function pushApplyDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], value: unknown): void {
  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', 'apply', 'Workbook run result description apply must be a plain object'))
    return
  }
  pushUnexpectedObjectFields(
    issues,
    value,
    'apply',
    new Set(['matched', 'planId', 'baseRevision', 'revision', 'previewOps', 'appliedOps', 'commandReceipts', 'proof']),
  )
  const matched = ownRunDescriptionDataValue(value, 'matched')
  if (matched !== true && matched !== false && matched !== null) {
    issues.push(
      runResultDescriptionIssue(
        'invalid_field',
        'apply.matched',
        'Workbook run result description apply.matched must be true, false, or null',
      ),
    )
  }
  pushOptionalStringFieldIssue(issues, value, 'apply', 'planId')
  pushSafeNonNegativeIntegerFieldIssue(issues, value, 'apply', 'baseRevision')
  pushSafeNonNegativeIntegerFieldIssue(issues, value, 'apply', 'revision')
  pushJsonArrayFieldIssue(issues, value, 'apply', 'previewOps')
  pushJsonArrayFieldIssue(issues, value, 'apply', 'appliedOps')
  pushJsonArrayFieldIssue(issues, value, 'apply', 'commandReceipts')
  pushRunDescriptionArrayEntryIssues(issues, value, 'commandReceipts', (entry, entryPath) =>
    pushCommandReceiptDescriptionIssues(issues, entry, `apply.${entryPath}`),
  )
}

function pushUndoDescriptionIssues(issues: WorkbookRunResultDescriptionIssue[], value: unknown): void {
  if (!isJsonDescriptionObject(value)) {
    issues.push(runResultDescriptionIssue('invalid_type', 'undo', 'Workbook run result description undo must be a plain object'))
    return
  }
  pushUnexpectedObjectFields(issues, value, 'undo', new Set(['id', 'ops']))
  pushRequiredStringFieldIssue(issues, value, 'undo', 'id')
  pushJsonArrayFieldIssue(issues, value, 'undo', 'ops')
}

function dedupeRunResultDescriptionIssues(
  issues: readonly WorkbookRunResultDescriptionIssue[],
): readonly WorkbookRunResultDescriptionIssue[] {
  const seen = new Set<string>()
  const deduped: WorkbookRunResultDescriptionIssue[] = []
  for (const issue of issues) {
    const key = `${issue.code}\u0000${issue.path}\u0000${issue.message}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    deduped.push(issue)
  }
  return deduped
}

function pushRequiredRunDescriptionArrayIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  key: string,
): void {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(runResultDescriptionIssue('missing_field', key, `Workbook run result description ${key} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(runResultDescriptionIssue('invalid_field', key, `Workbook run result description ${key} must be a data property`))
    return
  }
  if (!Array.isArray(descriptor.value)) {
    issues.push(runResultDescriptionIssue('invalid_field', key, `Workbook run result description ${key} must be an array`))
    return
  }
  pushJsonDescriptionValueIssues(issues, descriptor.value, key)
}

function pushOptionalRunDescriptionValueIssue(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  key: string,
): void {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(runResultDescriptionIssue('invalid_field', key, `Workbook run result description ${key} must be a data property`))
    return
  }
  pushJsonDescriptionValueIssues(issues, descriptor.value, key)
}

function isWorkbookRunResultDescriptionCore(value: unknown): value is WorkbookRunResultDescription {
  if (!isJsonDescriptionObject(value)) {
    return false
  }
  const status = ownJsonDescriptionValue(value, 'status')
  const common =
    hasJsonDescriptionArray(value, 'changed') &&
    hasJsonDescriptionArray(value, 'checks') &&
    hasOptionalJsonDescriptionValue(value, 'apply') &&
    hasOptionalJsonDescriptionValue(value, 'undo') &&
    hasOptionalJsonDescriptionValue(value, 'unverified')

  if (status === 'done') {
    return common && Object.getOwnPropertyDescriptor(value, 'errors') === undefined
  }
  if (status === 'failed') {
    return common && hasJsonDescriptionArray(value, 'errors')
  }
  return false
}

export function checkWorkbookRunResultDescription(value: unknown): WorkbookRunResultDescriptionCheckResult {
  if (!isJsonDescriptionObject(value)) {
    return freezeResultDescriptionData({
      status: 'invalid',
      issues: [runResultDescriptionIssue('invalid_type', 'result', 'Workbook run result description must be a plain object')],
    })
  }

  const issues: WorkbookRunResultDescriptionIssue[] = []
  const statusDescriptor = Object.getOwnPropertyDescriptor(value, 'status')
  let status: unknown
  if (statusDescriptor === undefined) {
    issues.push(runResultDescriptionIssue('missing_field', 'status', 'Workbook run result description status is required'))
  } else if (!('value' in statusDescriptor)) {
    issues.push(runResultDescriptionIssue('invalid_field', 'status', 'Workbook run result description status must be a data property'))
  } else {
    status = statusDescriptor.value
    if (status !== 'done' && status !== 'failed') {
      issues.push(runResultDescriptionIssue('invalid_status', 'status', 'Workbook run result description status must be done or failed'))
    }
  }

  if (status === 'done') {
    pushUnexpectedRunDescriptionFields(issues, value, new Set(['status', 'apply', 'changed', 'checks', 'undo', 'unverified']))
    pushRequiredRunDescriptionArrayIssue(issues, value, 'changed')
    pushRequiredRunDescriptionArrayIssue(issues, value, 'checks')
    pushOptionalRunDescriptionValueIssue(issues, value, 'apply')
    pushOptionalRunDescriptionValueIssue(issues, value, 'undo')
    pushOptionalRunDescriptionValueIssue(issues, value, 'unverified')
    pushRunDescriptionArrayEntryIssues(issues, value, 'changed', (entry, entryPath) =>
      pushChangeDescriptionIssues(issues, entry, entryPath),
    )
    pushRunDescriptionArrayEntryIssues(issues, value, 'checks', (entry, entryPath) => pushCheckDescriptionIssues(issues, entry, entryPath))
    pushRunDescriptionArrayEntryIssues(issues, value, 'unverified', (entry, entryPath) =>
      pushUnverifiedDescriptionIssues(issues, entry, entryPath),
    )
  } else if (status === 'failed') {
    pushUnexpectedRunDescriptionFields(issues, value, new Set(['status', 'errors', 'apply', 'changed', 'checks', 'undo', 'unverified']))
    pushRequiredRunDescriptionArrayIssue(issues, value, 'errors')
    pushRequiredRunDescriptionArrayIssue(issues, value, 'changed')
    pushRequiredRunDescriptionArrayIssue(issues, value, 'checks')
    pushOptionalRunDescriptionValueIssue(issues, value, 'apply')
    pushOptionalRunDescriptionValueIssue(issues, value, 'undo')
    pushOptionalRunDescriptionValueIssue(issues, value, 'unverified')
    pushRunDescriptionArrayEntryIssues(issues, value, 'errors', (entry, entryPath) => pushErrorDescriptionIssues(issues, entry, entryPath))
    pushRunDescriptionArrayEntryIssues(issues, value, 'changed', (entry, entryPath) =>
      pushChangeDescriptionIssues(issues, entry, entryPath),
    )
    pushRunDescriptionArrayEntryIssues(issues, value, 'checks', (entry, entryPath) => pushCheckDescriptionIssues(issues, entry, entryPath))
    pushRunDescriptionArrayEntryIssues(issues, value, 'unverified', (entry, entryPath) =>
      pushUnverifiedDescriptionIssues(issues, entry, entryPath),
    )
  }

  const apply = ownRunDescriptionDataValue(value, 'apply')
  if (apply !== undefined) {
    pushApplyDescriptionIssues(issues, apply)
  }
  const undo = ownRunDescriptionDataValue(value, 'undo')
  if (undo !== undefined) {
    pushUndoDescriptionIssues(issues, undo)
  }

  const dedupedIssues = dedupeRunResultDescriptionIssues(issues)
  if (dedupedIssues.length > 0 || !isWorkbookRunResultDescriptionCore(value)) {
    return freezeResultDescriptionData({
      status: 'invalid',
      issues:
        dedupedIssues.length > 0
          ? dedupedIssues
          : [runResultDescriptionIssue('invalid_field', 'result', 'Workbook run result description is invalid')],
    })
  }

  return freezeResultDescriptionData({
    status: 'valid',
    description: cloneResultDescriptionData(value),
    issues: [],
  })
}

export function isWorkbookRunResultDescription(value: unknown): value is WorkbookRunResultDescription {
  return checkWorkbookRunResultDescription(value).status === 'valid'
}
