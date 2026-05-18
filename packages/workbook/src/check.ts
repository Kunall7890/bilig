import type { WorkbookRef } from './find.js'
import type { WorkbookCheckResult } from './result.js'

export interface WorkbookCheckApi {
  readonly exists: (target: WorkbookRef) => WorkbookCheckResult
  readonly noFormulaErrors: (target: WorkbookRef) => WorkbookCheckResult
}

export function createWorkbookCheckResult(kind: string, target: WorkbookRef, message: string): WorkbookCheckResult {
  return {
    status: 'planned',
    kind,
    target,
    message,
  }
}

export function createWorkbookCheckApi(record?: (check: WorkbookCheckResult) => void): WorkbookCheckApi {
  function planned(kind: string, target: WorkbookRef, message: string): WorkbookCheckResult {
    const check = createWorkbookCheckResult(kind, target, message)
    record?.(check)
    return check
  }

  return {
    exists(target) {
      return planned('exists', target, `${target.label} exists`)
    },
    noFormulaErrors(target) {
      return planned('noFormulaErrors', target, `${target.label} has no formula errors`)
    },
  }
}

export const check: WorkbookCheckApi = createWorkbookCheckApi()
