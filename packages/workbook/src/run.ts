import type { WorkbookRef } from './find.js'
import type { WorkbookActionInput } from './input.js'
import { planWorkbookAction, type WorkbookActionMap, type WorkbookActionPlan, type WorkbookModel } from './model.js'
import type { EngineOp } from './ops.js'
import {
  checkPlanData,
  hydratePlanData,
  isHydratedPlan,
  type WorkbookExecutablePlan,
  type WorkbookPlanId,
  type WorkbookPlanDataRefs,
} from './plan-data.js'
import { verifyWorkbookReadbacks, type WorkbookRunReadback } from './readback.js'
import { checkRuntimeAdapter, type WorkbookRuntimeCapability } from './requirements.js'
import {
  applyUnverified,
  changedAfterApply,
  changedAfterOptionalApply,
  describeApply,
  failedApplyResult,
  validateApplyResult,
} from './run-apply.js'
import { unverifiedCheckErrors, verifyChecksWithAdapter } from './run-check-verification.js'
import { errorMessage, failedRun, planDataRunError, runError } from './run-failure.js'
import {
  adapterApplyMethod,
  adapterReadMethod,
  applyProofErrors,
  checkProofErrors,
  normalizeRunOptions,
  preApplyProofErrors,
} from './run-runtime-boundary.js'
import { freezeWorkbookRunResult } from './run-result.js'
import type {
  WorkbookCheckResult,
  WorkbookRunApplyCommandReceipt,
  WorkbookRunApplySummary,
  WorkbookRunError,
  WorkbookRunResult,
  WorkbookRunUnverified,
  WorkbookUndoRef,
} from './result.js'
import { verifyPlan } from './verify.js'

type MaybePromise<T> = T | Promise<T>

export interface WorkbookRunApplyResult {
  readonly status: 'applied' | 'failed'
  readonly planId?: WorkbookPlanId
  readonly baseRevision?: number
  readonly revision?: number
  readonly previewOps?: readonly EngineOp[]
  readonly appliedOps?: readonly EngineOp[]
  readonly commandReceipts?: readonly WorkbookRunApplyCommandReceipt[]
  readonly proof?: WorkbookActionInput
  readonly errors?: readonly WorkbookRunError[]
  readonly undo?: WorkbookUndoRef
}

export interface WorkbookRunOptions {
  readonly strict?: boolean
  readonly requireApplyProof?: boolean
  readonly requirePlanId?: boolean
  readonly requireResolvedRefs?: boolean
  readonly requireChecks?: boolean
  readonly requireCheckProof?: boolean
  readonly requireRevision?: boolean
  readonly requireNoUnverified?: boolean
  readonly expectedBaseRevision?: number
}

export interface WorkbookRunAdapter<Refs = unknown> {
  apply?(plan: WorkbookActionPlan<Refs>): MaybePromise<WorkbookRunApplyResult>
  read?(targets: readonly WorkbookRef[], plan: WorkbookActionPlan<Refs>): MaybePromise<readonly WorkbookRunReadback[]>
  verifyChecks?(checks: readonly WorkbookCheckResult[], plan: WorkbookActionPlan<Refs>): MaybePromise<readonly WorkbookCheckResult[]>
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

  return failedRun({
    errors: verification.issues.map((issue) => runError(issue.code, issue.message)),
    checks: plan.checks,
  })
}

function unverifiedProperty(unverified: readonly WorkbookRunUnverified[]): { readonly unverified: readonly WorkbookRunUnverified[] } | {} {
  return unverified.length === 0 ? {} : { unverified }
}

function needsApply(capabilities: readonly WorkbookRuntimeCapability[]): boolean {
  return capabilities.some(
    (capability) =>
      capability === 'writeFormula' ||
      capability === 'writeValue' ||
      capability === 'format' ||
      capability === 'clear' ||
      capability === 'applyOp',
  )
}

async function runLiveWorkbookPlan<Refs>(
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
  options: unknown = {},
): Promise<WorkbookRunResult> {
  const invalidPlan = failedFromPlanIssues(plan)
  if (invalidPlan !== null) {
    return invalidPlan
  }

  const normalizedOptions = normalizeRunOptions(options)
  if (normalizedOptions.status === 'invalid') {
    return failedRun({
      errors: [normalizedOptions.error],
      checks: plan.checks,
    })
  }

  const adapterCheck = checkRuntimeAdapter(plan, adapter)
  if (adapterCheck.status === 'invalid') {
    return failedRun({
      errors: adapterCheck.issues.map((adapterIssue) => runError('adapter_missing_capability', adapterIssue.message)),
      checks: plan.checks,
    })
  }

  let validApplyResult: WorkbookRunApplyResult | undefined
  let apply: WorkbookRunApplySummary | undefined
  let unverified: readonly WorkbookRunUnverified[] = []
  const requiresApply = needsApply(adapterCheck.requiredCapabilities)
  const preApplyErrors = preApplyProofErrors(plan, requiresApply, normalizedOptions.options)
  if (preApplyErrors.length > 0) {
    return failedRun({
      errors: preApplyErrors,
      checks: plan.checks,
    })
  }
  if (requiresApply) {
    const applyMethod = adapterApplyMethod(adapter)
    let applyResult: unknown
    try {
      applyResult = applyMethod === undefined ? undefined : await Reflect.apply(applyMethod, adapter, [plan])
    } catch (error) {
      return failedRun({
        errors: [runError('apply_failed', errorMessage(error))],
        checks: plan.checks,
      })
    }

    const applyValidation = validateApplyResult(plan, applyResult)
    if (applyValidation.status === 'invalid') {
      return applyValidation.result
    }
    validApplyResult = applyValidation.result

    if (validApplyResult.status === 'failed') {
      return failedApplyResult(plan, validApplyResult)
    }

    apply = describeApply(validApplyResult)
    unverified = applyUnverified(plan, apply)
    const applyErrors = applyProofErrors(plan, apply, normalizedOptions.options)
    if (applyErrors.length > 0) {
      return failedRun({
        errors: applyErrors,
        apply,
        changed: changedAfterApply(plan, validApplyResult),
        checks: plan.checks,
        ...(validApplyResult.undo !== undefined ? { undo: validApplyResult.undo } : {}),
        unverified,
      })
    }
  }

  let checks = plan.checks
  const targets = readbackTargets(checks)
  if (targets.length > 0) {
    const read = adapterReadMethod(adapter)
    if (read === undefined) {
      const readbackVerification = verifyWorkbookReadbacks(checks, [])
      return failedRun({
        errors: readbackVerification.issues.map((issue) => runError(issue.code, issue.message)),
        apply,
        changed: changedAfterOptionalApply(plan, validApplyResult),
        checks: readbackVerification.checks,
        ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
        unverified,
      })
    }

    let readbacks: readonly WorkbookRunReadback[]
    try {
      readbacks = await Reflect.apply(read, adapter, [targets, plan])
    } catch (error) {
      return failedRun({
        errors: [runError('readback_failed', errorMessage(error))],
        apply,
        changed: changedAfterOptionalApply(plan, validApplyResult),
        checks,
        ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
        unverified,
      })
    }

    const readbackVerification = verifyWorkbookReadbacks(checks, readbacks)
    checks = readbackVerification.checks
    if (readbackVerification.status === 'failed') {
      return failedRun({
        errors: readbackVerification.issues.map((issue) => runError(issue.code, issue.message)),
        apply,
        changed: changedAfterOptionalApply(plan, validApplyResult),
        checks,
        ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
        unverified,
      })
    }
  }

  const checkVerification = await verifyChecksWithAdapter(checks, plan, adapter)
  checks = checkVerification.checks
  if (checkVerification.errors.length > 0) {
    return failedRun({
      errors: checkVerification.errors,
      apply,
      changed: changedAfterOptionalApply(plan, validApplyResult),
      checks,
      ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
      unverified,
    })
  }

  const proofErrors = checkProofErrors(checks, normalizedOptions.options)
  if (proofErrors.length > 0) {
    return failedRun({
      errors: proofErrors,
      apply,
      changed: changedAfterOptionalApply(plan, validApplyResult),
      checks,
      ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
      unverified,
    })
  }

  const unverifiedErrors = unverifiedCheckErrors(checks)
  if (unverifiedErrors.length > 0) {
    return failedRun({
      errors: unverifiedErrors,
      apply,
      changed: changedAfterOptionalApply(plan, validApplyResult),
      checks,
      ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
      unverified,
    })
  }

  return {
    status: 'done',
    ...(apply !== undefined ? { apply } : {}),
    changed: changedAfterOptionalApply(plan, validApplyResult),
    checks,
    ...(validApplyResult?.undo !== undefined ? { undo: validApplyResult.undo } : {}),
    ...unverifiedProperty(unverified),
  }
}

export function runWorkbookPlan<Refs>(
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
  options?: WorkbookRunOptions,
): Promise<WorkbookRunResult>
export function runWorkbookPlan(
  plan: WorkbookExecutablePlan,
  adapter: WorkbookRunAdapter<WorkbookPlanDataRefs>,
  options?: WorkbookRunOptions,
): Promise<WorkbookRunResult>
export function runWorkbookPlan(plan: unknown, adapter: WorkbookRunAdapter, options?: WorkbookRunOptions): Promise<WorkbookRunResult>
export function runWorkbookPlan(plan: unknown, adapter: WorkbookRunAdapter, options: unknown = {}): Promise<WorkbookRunResult> {
  if (isHydratedPlan(plan)) {
    return runLiveWorkbookPlan<unknown>(plan, adapter, options).then(freezeWorkbookRunResult)
  }
  const planDataCheck = checkPlanData(plan)
  if (planDataCheck.status === 'invalid') {
    return Promise.resolve(
      freezeWorkbookRunResult(
        failedRun({
          errors: planDataCheck.issues.map(planDataRunError),
          checks: [],
        }),
      ),
    )
  }
  return runLiveWorkbookPlan<unknown>(hydratePlanData(planDataCheck.plan), adapter, options).then(freezeWorkbookRunResult)
}

export async function runWorkbookAction<Refs, Actions extends WorkbookActionMap<Refs>>(
  model: WorkbookModel<Refs, Actions>,
  actionName: string,
  adapter: WorkbookRunAdapter<Refs>,
  input?: WorkbookActionInput,
  options?: WorkbookRunOptions,
): Promise<WorkbookRunResult> {
  const result = planWorkbookAction(model, actionName, input)
  if (result.status === 'failed') {
    return freezeWorkbookRunResult(
      failedRun({
        errors: result.errors,
        checks: result.checks,
      }),
    )
  }
  return freezeWorkbookRunResult(await runLiveWorkbookPlan(result.plan, adapter, options))
}
