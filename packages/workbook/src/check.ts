import type { WorkbookRef } from './find.js'
import type { WorkbookCheckResult, WorkbookCheckStatus } from './result.js'

export interface WorkbookCustomCheckOptions {
  readonly kind: string
  readonly message: string
  readonly target?: WorkbookRef
  readonly refs?: readonly WorkbookRef[]
  readonly status?: WorkbookCheckStatus
}

export interface WorkbookCheckApi {
  readonly exists: (target: WorkbookRef) => WorkbookCheckResult
  readonly noFormulaErrors: (target: WorkbookRef) => WorkbookCheckResult
  readonly custom: (options: WorkbookCustomCheckOptions) => WorkbookCheckResult
}

export function createWorkbookCheckResult(kind: string, target: WorkbookRef, message: string): WorkbookCheckResult {
  return createWorkbookCustomCheck({ kind, target, message })
}

function requiredText(value: string, name: string): string {
  const trimmed = value.trim()
  if (trimmed === '') {
    throw new Error(`Workbook check ${name} cannot be empty`)
  }
  return trimmed
}

function normalizeCheckStatus(status: WorkbookCheckStatus | undefined): WorkbookCheckStatus {
  if (status === undefined || status === 'planned' || status === 'passed' || status === 'failed') {
    return status ?? 'planned'
  }
  throw new Error(`Unsupported workbook check status: ${String(status)}`)
}

function refKey(ref: WorkbookRef): string {
  return `${ref.kind}:${ref.id}`
}

function uniqueRefs(refs: readonly WorkbookRef[] | undefined): readonly WorkbookRef[] | undefined {
  if (refs === undefined) {
    return undefined
  }
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  for (const ref of refs) {
    const key = refKey(ref)
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(ref)
  }
  return unique.length === 0 ? undefined : unique
}

export function createWorkbookCustomCheck(options: WorkbookCustomCheckOptions): WorkbookCheckResult {
  const refs = uniqueRefs(options.refs)
  return {
    status: normalizeCheckStatus(options.status),
    kind: requiredText(options.kind, 'kind'),
    ...(options.target !== undefined ? { target: options.target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message: requiredText(options.message, 'message'),
  }
}

export function createWorkbookCheckApi(record?: (check: WorkbookCheckResult) => void): WorkbookCheckApi {
  function planned(options: WorkbookCustomCheckOptions): WorkbookCheckResult {
    const check = createWorkbookCustomCheck(options)
    record?.(check)
    return check
  }

  return {
    exists(target) {
      return planned({ kind: 'exists', target, message: `${target.label} exists` })
    },
    noFormulaErrors(target) {
      return planned({ kind: 'noFormulaErrors', target, message: `${target.label} has no formula errors` })
    },
    custom(options) {
      return planned(options)
    },
  }
}

export const check: WorkbookCheckApi = createWorkbookCheckApi()
