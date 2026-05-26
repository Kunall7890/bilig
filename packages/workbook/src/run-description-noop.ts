import type { WorkbookRunResultDescriptionIssue } from './run-description.js'

function issue(code: WorkbookRunResultDescriptionIssue['code'], path: string, message: string): WorkbookRunResultDescriptionIssue {
  return {
    code,
    path,
    message,
  }
}

function ownValue(value: Record<string, unknown>, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  return descriptor !== undefined && 'value' in descriptor ? descriptor.value : undefined
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const prototype = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function pushUnexpectedFields(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  expected: ReadonlySet<string>,
): void {
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !expected.has(key)) {
      issues.push(
        issue('unexpected_field', `${path}.${String(key)}`, `Workbook run result description ${path} must not include ${String(key)}`),
      )
    }
  }
}

function pushJsonValueIssues(
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
      issues.push(issue('invalid_field', path, `Workbook run result description ${path} must be JSON-safe`))
    }
    return
  }
  if (typeof value !== 'object') {
    issues.push(issue('invalid_field', path, `Workbook run result description ${path} must be JSON-safe`))
    return
  }
  if (seen.has(value)) {
    issues.push(issue('invalid_field', path, `Workbook run result description ${path} must not be circular`))
    return
  }
  seen.add(value)

  if (Array.isArray(value)) {
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const entryPath = `${path}[${index}]`
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor)) {
        issues.push(issue('invalid_field', entryPath, `Workbook run result description ${entryPath} must be a data property`))
        continue
      }
      pushJsonValueIssues(issues, descriptor.value, entryPath, seen)
    }
    return
  }

  if (!isPlainObject(value)) {
    issues.push(issue('invalid_type', path, `Workbook run result description ${path} must be a plain object`))
    return
  }
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (!descriptor.enumerable) {
      continue
    }
    const entryPath = `${path}.${key}`
    if (!('value' in descriptor)) {
      issues.push(issue('invalid_field', entryPath, `Workbook run result description ${entryPath} must be a data property`))
      continue
    }
    pushJsonValueIssues(issues, descriptor.value, entryPath, seen)
  }
}

function pushRequiredTrimmedString(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    issues.push(issue('missing_field', fieldPath, `Workbook run result description ${fieldPath} is required`))
    return
  }
  if (!('value' in descriptor)) {
    issues.push(issue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`))
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(issue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a string`))
    return
  }
  if (descriptor.value.trim() === '' || descriptor.value.trim() !== descriptor.value) {
    issues.push(
      issue(
        'invalid_field',
        fieldPath,
        `Workbook run result description ${fieldPath} must not be empty or have leading or trailing whitespace`,
      ),
    )
  }
}

function pushMatchingString(
  issues: WorkbookRunResultDescriptionIssue[],
  value: Record<string, unknown>,
  path: string,
  key: string,
  expected: unknown,
  expectedLabel: string,
): void {
  const actual = ownValue(value, key)
  if (typeof actual === 'string' && typeof expected === 'string' && actual !== expected) {
    issues.push(issue('invalid_field', `${path}.${key}`, `Workbook run result description ${path}.${key} must match ${expectedLabel}`))
  }
}

function pushOptionalString(issues: WorkbookRunResultDescriptionIssue[], value: Record<string, unknown>, path: string, key: string): void {
  const fieldPath = `${path}.${key}`
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined) {
    return
  }
  if (!('value' in descriptor)) {
    issues.push(issue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a data property`))
    return
  }
  if (typeof descriptor.value !== 'string') {
    issues.push(issue('invalid_field', fieldPath, `Workbook run result description ${fieldPath} must be a string`))
  }
}

function pushNoopOpsIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  entry: Record<string, unknown>,
  path: string,
  key: 'previewOps' | 'appliedOps',
): void {
  const ops = ownValue(entry, key)
  if (Array.isArray(ops) && ops.length > 0) {
    issues.push(
      issue('invalid_field', `${path}.${key}`, `Workbook run result description ${path}.${key} must be empty when noop is present`),
    )
  }
}

function pushOpEffectIssues(issues: WorkbookRunResultDescriptionIssue[], effect: Record<string, unknown>, path: string): void {
  const opKind = Object.getOwnPropertyDescriptor(effect, 'opKind')
  if (opKind === undefined) {
    issues.push(issue('missing_field', `${path}.opKind`, `Workbook run result description ${path}.opKind is required`))
  } else if (!('value' in opKind)) {
    issues.push(issue('invalid_field', `${path}.opKind`, `Workbook run result description ${path}.opKind must be a data property`))
  } else if (typeof opKind.value !== 'string' || opKind.value.trim() === '' || opKind.value.trim() !== opKind.value) {
    issues.push(
      issue(
        'invalid_field',
        `${path}.opKind`,
        `Workbook run result description ${path}.opKind must not be empty or have leading or trailing whitespace`,
      ),
    )
  }

  const op = Object.getOwnPropertyDescriptor(effect, 'op')
  if (op === undefined) {
    issues.push(issue('missing_field', `${path}.op`, `Workbook run result description ${path}.op is required`))
  } else if (!('value' in op)) {
    issues.push(issue('invalid_field', `${path}.op`, `Workbook run result description ${path}.op must be a data property`))
  } else if (!isPlainObject(op.value)) {
    issues.push(issue('invalid_type', `${path}.op`, `Workbook run result description ${path}.op must be a plain object`))
  }
}

function pushEffectIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  effect: Record<string, unknown>,
  path: string,
  expectedCommandKind: unknown,
): void {
  const kind = Object.getOwnPropertyDescriptor(effect, 'kind')
  if (kind === undefined) {
    issues.push(issue('missing_field', `${path}.kind`, `Workbook run result description ${path}.kind is required`))
    return
  }
  if (!('value' in kind)) {
    issues.push(issue('invalid_field', `${path}.kind`, `Workbook run result description ${path}.kind must be a data property`))
    return
  }
  if (typeof kind.value !== 'string') {
    issues.push(issue('invalid_field', `${path}.kind`, `Workbook run result description ${path}.kind must be a string`))
    return
  }
  if (typeof expectedCommandKind === 'string' && kind.value !== expectedCommandKind) {
    issues.push(issue('invalid_field', `${path}.kind`, `Workbook run result description ${path}.kind must match receipt commandKind`))
  }
  if (expectedCommandKind === 'op') {
    pushOpEffectIssues(issues, effect, path)
  }
}

function pushNoopProofIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  proof: Record<string, unknown>,
  path: string,
  expected: { readonly commandKind: unknown; readonly commandDigest: unknown },
): void {
  pushRequiredTrimmedString(issues, proof, path, 'source')
  pushRequiredTrimmedString(issues, proof, path, 'evidence')
  pushRequiredTrimmedString(issues, proof, path, 'commandKind')
  pushRequiredTrimmedString(issues, proof, path, 'commandDigest')
  pushMatchingString(issues, proof, path, 'commandKind', expected.commandKind, 'receipt commandKind')
  pushMatchingString(issues, proof, path, 'commandDigest', expected.commandDigest, 'receipt commandDigest')

  const opCount = Object.getOwnPropertyDescriptor(proof, 'opCount')
  if (opCount === undefined) {
    issues.push(issue('missing_field', `${path}.opCount`, `Workbook run result description ${path}.opCount is required`))
  } else if (!('value' in opCount)) {
    issues.push(issue('invalid_field', `${path}.opCount`, `Workbook run result description ${path}.opCount must be a data property`))
  } else if (opCount.value !== 0) {
    issues.push(issue('invalid_field', `${path}.opCount`, `Workbook run result description ${path}.opCount must be 0`))
  }

  const effect = Object.getOwnPropertyDescriptor(proof, 'effect')
  if (effect === undefined) {
    issues.push(issue('missing_field', `${path}.effect`, `Workbook run result description ${path}.effect is required`))
    return
  }
  if (!('value' in effect)) {
    issues.push(issue('invalid_field', `${path}.effect`, `Workbook run result description ${path}.effect must be a data property`))
    return
  }
  if (!isPlainObject(effect.value)) {
    issues.push(issue('invalid_type', `${path}.effect`, `Workbook run result description ${path}.effect must be a plain object`))
    return
  }
  pushEffectIssues(issues, effect.value, `${path}.effect`, expected.commandKind)
  pushJsonValueIssues(issues, effect.value, `${path}.effect`)
}

function pushNoopIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  value: unknown,
  path: string,
  expected: { readonly commandKind: unknown; readonly commandDigest: unknown },
): void {
  if (!isPlainObject(value)) {
    issues.push(issue('invalid_type', path, `Workbook run result description ${path} must be a plain object`))
    return
  }
  pushUnexpectedFields(issues, value, path, new Set(['reason', 'message', 'proof']))

  const reason = Object.getOwnPropertyDescriptor(value, 'reason')
  if (reason === undefined) {
    issues.push(issue('missing_field', `${path}.reason`, `Workbook run result description ${path}.reason is required`))
  } else if (!('value' in reason)) {
    issues.push(issue('invalid_field', `${path}.reason`, `Workbook run result description ${path}.reason must be a data property`))
  } else if (reason.value !== 'already_satisfied') {
    issues.push(issue('invalid_field', `${path}.reason`, `Workbook run result description ${path}.reason must be already_satisfied`))
  }

  pushOptionalString(issues, value, path, 'message')

  const proof = Object.getOwnPropertyDescriptor(value, 'proof')
  if (proof === undefined) {
    issues.push(issue('missing_field', `${path}.proof`, `Workbook run result description ${path}.proof is required`))
    return
  }
  if (!('value' in proof)) {
    issues.push(issue('invalid_field', `${path}.proof`, `Workbook run result description ${path}.proof must be a data property`))
    return
  }
  if (!isPlainObject(proof.value)) {
    issues.push(issue('invalid_type', `${path}.proof`, `Workbook run result description ${path}.proof must be a plain object`))
    return
  }
  pushNoopProofIssues(issues, proof.value, `${path}.proof`, expected)
}

export function pushCommandReceiptNoopDescriptionIssues(
  issues: WorkbookRunResultDescriptionIssue[],
  entry: Record<string, unknown>,
  path: string,
): void {
  const noop = ownValue(entry, 'noop')
  if (noop === undefined) {
    return
  }
  pushNoopIssues(issues, noop, `${path}.noop`, {
    commandKind: ownValue(entry, 'commandKind'),
    commandDigest: ownValue(entry, 'commandDigest'),
  })
  pushNoopOpsIssues(issues, entry, path, 'previewOps')
  pushNoopOpsIssues(issues, entry, path, 'appliedOps')
}
