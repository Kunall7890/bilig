import type { WorkbookRef } from './find.js'
import type { WorkbookActionInput } from './input.js'
import { planWorkbookAction, type WorkbookActionMap, type WorkbookActionPlan, type WorkbookModel } from './model.js'
import { verifyWorkbookReadbacks, type WorkbookRunReadback } from './readback.js'
import type { WorkbookCheckResult, WorkbookRunError, WorkbookRunResult, WorkbookUndoRef } from './result.js'
import { verifyPlan } from './verify.js'

type MaybePromise<T> = T | Promise<T>

export interface WorkbookRunApplyResult {
  readonly status: 'applied' | 'failed'
  readonly errors?: readonly WorkbookRunError[]
  readonly checks?: readonly WorkbookCheckResult[]
  readonly undo?: WorkbookUndoRef
}

export interface WorkbookRunAdapter<Refs = unknown> {
  readonly apply: (plan: WorkbookActionPlan<Refs>) => MaybePromise<WorkbookRunApplyResult>
  readonly read?: (targets: readonly WorkbookRef[], plan: WorkbookActionPlan<Refs>) => MaybePromise<readonly WorkbookRunReadback[]>
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function runError(code: string, message: string): WorkbookRunError {
  return {
    code,
    message,
  }
}

function readbackTargets(checks: readonly WorkbookCheckResult[]): readonly WorkbookRef[] {
  const targets: WorkbookRef[] = []
  const seen = new Set<string>()
  checks.forEach((check) => {
    if (check.expectation === undefined || check.target === undefined) {
      return
    }
    const key = `${check.target.kind}:${check.target.id}`
    if (seen.has(key)) {
      return
    }
    seen.add(key)
    targets.push(check.target)
  })
  return targets
}

function failedFromPlanIssues<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRunResult | null {
  const verification = verifyPlan(plan)
  if (verification.status === 'valid') {
    return null
  }

  return {
    status: 'failed',
    errors: verification.issues.map((issue) => runError(issue.code, issue.message)),
    checks: plan.checks,
  }
}

function failedApplyResult(plan: WorkbookActionPlan, result: WorkbookRunApplyResult): WorkbookRunResult {
  return {
    status: 'failed',
    errors: result.errors ?? [runError('apply_failed', `Workbook action ${plan.modelName}.${plan.actionName} failed to apply`)],
    checks: result.checks ?? plan.checks,
  }
}

export async function runWorkbookPlan<Refs>(plan: WorkbookActionPlan<Refs>, adapter: WorkbookRunAdapter<Refs>): Promise<WorkbookRunResult> {
  const invalidPlan = failedFromPlanIssues(plan)
  if (invalidPlan !== null) {
    return invalidPlan
  }

  let applyResult: WorkbookRunApplyResult
  try {
    applyResult = await adapter.apply(plan)
  } catch (error) {
    return {
      status: 'failed',
      errors: [runError('apply_failed', errorMessage(error))],
      checks: plan.checks,
    }
  }

  if (applyResult.status === 'failed') {
    return failedApplyResult(plan, applyResult)
  }

  let checks = applyResult.checks ?? plan.checks
  const targets = readbackTargets(checks)
  if (targets.length > 0) {
    if (adapter.read === undefined) {
      const readbackVerification = verifyWorkbookReadbacks(checks, [])
      return {
        status: 'failed',
        errors: readbackVerification.issues.map((issue) => runError(issue.code, issue.message)),
        checks: readbackVerification.checks,
      }
    }

    let readbacks: readonly WorkbookRunReadback[]
    try {
      readbacks = await adapter.read(targets, plan)
    } catch (error) {
      return {
        status: 'failed',
        errors: [runError('readback_failed', errorMessage(error))],
        checks,
      }
    }

    const readbackVerification = verifyWorkbookReadbacks(checks, readbacks)
    checks = readbackVerification.checks
    if (readbackVerification.status === 'failed') {
      return {
        status: 'failed',
        errors: readbackVerification.issues.map((issue) => runError(issue.code, issue.message)),
        checks,
      }
    }
  }

  return {
    status: 'done',
    changed: plan.changed,
    checks,
    ...(applyResult.undo !== undefined ? { undo: applyResult.undo } : {}),
  }
}

export async function runWorkbookAction<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
  adapter: WorkbookRunAdapter<Refs>,
  input?: WorkbookActionInput,
): Promise<WorkbookRunResult> {
  const result = planWorkbookAction(model, actionName, input)
  if (result.status === 'failed') {
    return {
      status: 'failed',
      errors: result.errors,
      checks: result.checks,
    }
  }
  return runWorkbookPlan(result.plan, adapter)
}
