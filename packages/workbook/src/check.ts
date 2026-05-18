import type { WorkbookRef } from './find.js'
import type { WorkbookCheckResult, WorkbookCheckStatus } from './result.js'

export interface WorkbookCustomCheckOptions {
  readonly kind: string
  readonly message: string
  readonly target?: WorkbookRef
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

export function createWorkbookCustomCheck(options: WorkbookCustomCheckOptions): WorkbookCheckResult {
  return {
    status: normalizeCheckStatus(options.status),
    kind: requiredText(options.kind, 'kind'),
    ...(options.target !== undefined ? { target: options.target } : {}),
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
