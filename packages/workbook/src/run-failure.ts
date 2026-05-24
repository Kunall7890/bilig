import type { WorkbookPlanDataIssue } from './plan-data.js'
import type {
  WorkbookChangeSummary,
  WorkbookCheckResult,
  WorkbookRunApplySummary,
  WorkbookRunError,
  WorkbookRunErrorCode,
  WorkbookRunResult,
  WorkbookRunUnverified,
  WorkbookUndoRef,
} from './result.js'

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function runError(code: WorkbookRunErrorCode, message: string): WorkbookRunError {
  return {
    code,
    message,
  }
}

export function failedRun(args: {
  readonly errors: readonly WorkbookRunError[]
  readonly apply?: WorkbookRunApplySummary | undefined
  readonly changed?: readonly WorkbookChangeSummary[]
  readonly checks: readonly WorkbookCheckResult[]
  readonly undo?: WorkbookUndoRef
  readonly unverified?: readonly WorkbookRunUnverified[]
}): WorkbookRunResult {
  return {
    status: 'failed',
    errors: args.errors,
    ...(args.apply !== undefined ? { apply: args.apply } : {}),
    changed: args.changed ?? [],
    checks: args.checks,
    ...(args.undo !== undefined ? { undo: args.undo } : {}),
    ...(args.unverified !== undefined && args.unverified.length > 0 ? { unverified: args.unverified } : {}),
  }
}

export function planDataRunError(issue: WorkbookPlanDataIssue): WorkbookRunError {
  return {
    code: 'invalid_plan_data',
    message: issue.message,
    path: issue.path,
    issueCode: issue.code,
  }
}
