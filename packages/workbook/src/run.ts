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
  readonly verifyChecks?: (
    checks: readonly WorkbookCheckResult[],
    plan: WorkbookActionPlan<Refs>,
  ) => MaybePromise<readonly WorkbookCheckResult[]>
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

function checkLabel(check: WorkbookCheckResult): string {
  return check.target?.label ?? check.kind
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value))
}

function canonicalValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function checkContract(check: WorkbookCheckResult): Omit<WorkbookCheckResult, 'status'> {
  return {
    kind: check.kind,
    ...(check.target !== undefined ? { target: check.target } : {}),
    ...(check.refs !== undefined ? { refs: check.refs } : {}),
    message: check.message,
    ...(check.expectation !== undefined ? { expectation: check.expectation } : {}),
  }
}

function cloneCheck(check: WorkbookCheckResult): WorkbookCheckResult {
  return {
    status: check.status,
    kind: check.kind,
    ...(check.target !== undefined ? { target: check.target } : {}),
    ...(check.refs !== undefined ? { refs: check.refs } : {}),
    message: check.message,
    ...(check.expectation !== undefined ? { expectation: check.expectation } : {}),
  }
}

function checkContractMatches(expectedContract: string, actual: WorkbookCheckResult): boolean {
  return expectedContract === canonicalJson(checkContract(actual))
}

function isCheckStatus(value: unknown): value is WorkbookCheckResult['status'] {
  return value === 'planned' || value === 'passed' || value === 'failed'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorkbookCheckResult(value: unknown): value is WorkbookCheckResult {
  return isRecord(value) && isCheckStatus(value['status']) && typeof value['kind'] === 'string' && typeof value['message'] === 'string'
}

type CheckValidation =
  | {
      readonly status: 'valid'
      readonly checks: readonly WorkbookCheckResult[]
    }
  | {
      readonly status: 'invalid'
      readonly error: WorkbookRunError
    }

function validateVerifiedChecks(
  originalContracts: readonly string[],
  originalKinds: readonly string[],
  verified: unknown,
): CheckValidation {
  if (!Array.isArray(verified)) {
    return {
      status: 'invalid',
      error: runError('invalid_check_verification', 'Check verifier did not return a check array'),
    }
  }

  if (verified.length !== originalContracts.length) {
    return {
      status: 'invalid',
      error: runError(
        'invalid_check_verification',
        `Check verifier returned ${String(verified.length)} checks for ${String(originalContracts.length)} planned checks`,
      ),
    }
  }

  const verifiedChecks: WorkbookCheckResult[] = []
  for (let index = 0; index < originalContracts.length; index += 1) {
    const expectedContract = originalContracts[index]
    const expectedKind = originalKinds[index] ?? 'check'
    const actual = verified[index]
    if (expectedContract === undefined || !isRecord(actual)) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid check at index ${String(index)}`),
      }
    }
    if (!isCheckStatus(actual['status'])) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid status at index ${String(index)}`),
      }
    }
    if (!isWorkbookCheckResult(actual)) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid check at index ${String(index)}`),
      }
    }
    if (!checkContractMatches(expectedContract, actual)) {
      return {
        status: 'invalid',
        error: runError(
          'invalid_check_verification',
          `Check verifier changed the check contract at index ${String(index)} for ${expectedKind}`,
        ),
      }
    }
    verifiedChecks.push(actual)
  }

  return {
    status: 'valid',
    checks: verifiedChecks,
  }
}

async function verifyChecksWithAdapter<Refs>(
  checks: readonly WorkbookCheckResult[],
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
): Promise<{ readonly checks: readonly WorkbookCheckResult[]; readonly errors: readonly WorkbookRunError[] }> {
  if (adapter.verifyChecks === undefined) {
    return { checks, errors: [] }
  }

  const originalChecks = checks.map(cloneCheck)
  const originalContracts = originalChecks.map((check) => canonicalJson(checkContract(check)))
  const originalKinds = originalChecks.map((check) => check.kind)
  const verifierInput = originalChecks.map(cloneCheck)
  let verified: unknown
  try {
    verified = await adapter.verifyChecks(verifierInput, plan)
  } catch (error) {
    return {
      checks: originalChecks,
      errors: [runError('check_verification_failed', errorMessage(error))],
    }
  }

  const validation = validateVerifiedChecks(originalContracts, originalKinds, verified)
  if (validation.status === 'invalid') {
    return {
      checks: originalChecks,
      errors: [validation.error],
    }
  }

  const verifiedChecks = validation.checks
  const failedChecks = verifiedChecks.filter((check) => check.status === 'failed')
  if (failedChecks.length > 0) {
    return {
      checks: verifiedChecks,
      errors: failedChecks.map((check) =>
        runError('check_failed', `${check.target?.label ?? check.kind} failed check ${check.kind}: ${check.message}`),
      ),
    }
  }

  return { checks: verifiedChecks, errors: [] }
}

function unverifiedCheckErrors(checks: readonly WorkbookCheckResult[]): readonly WorkbookRunError[] {
  return checks
    .filter((check) => check.status === 'planned')
    .map((check) => runError('check_not_verified', `${checkLabel(check)} did not verify check ${check.kind}: ${check.message}`))
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

  const checkVerification = await verifyChecksWithAdapter(checks, plan, adapter)
  checks = checkVerification.checks
  if (checkVerification.errors.length > 0) {
    return {
      status: 'failed',
      errors: checkVerification.errors,
      checks,
    }
  }

  const unverifiedErrors = unverifiedCheckErrors(checks)
  if (unverifiedErrors.length > 0) {
    return {
      status: 'failed',
      errors: unverifiedErrors,
      checks,
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
