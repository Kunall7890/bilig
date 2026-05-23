import type { WorkbookRef } from './find.js'
import { isWorkbookOp } from './guards.js'
import { normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import { planWorkbookAction, type WorkbookActionMap, type WorkbookActionPlan, type WorkbookModel } from './model.js'
import type { EngineOp } from './ops.js'
import { hydratePlanData, isHydratedPlan, type WorkbookExecutablePlan, type WorkbookPlanDataRefs } from './plan-data.js'
import { verifyWorkbookReadbacks, type WorkbookRunReadback } from './readback.js'
import { checkRuntimeAdapter, type WorkbookRuntimeCapability } from './requirements.js'
import {
  isWorkbookRunErrorCode,
  type WorkbookChangeSummary,
  type WorkbookRunApplySummary,
  type WorkbookCheckResult,
  type WorkbookRunError,
  type WorkbookRunErrorCode,
  type WorkbookRunResult,
  type WorkbookRunUnverified,
  type WorkbookUndoRef,
} from './result.js'
import { verifyPlan } from './verify.js'

type MaybePromise<T> = T | Promise<T>

export interface WorkbookRunApplyResult {
  readonly status: 'applied' | 'failed'
  readonly previewOps?: readonly EngineOp[]
  readonly appliedOps?: readonly EngineOp[]
  readonly proof?: WorkbookActionInput
  readonly errors?: readonly WorkbookRunError[]
  readonly undo?: WorkbookUndoRef
}

export interface WorkbookRunOptions {
  readonly requireApplyProof?: boolean
}

export interface WorkbookRunAdapter<Refs = unknown> {
  apply?(plan: WorkbookActionPlan<Refs>): MaybePromise<WorkbookRunApplyResult>
  read?(targets: readonly WorkbookRef[], plan: WorkbookActionPlan<Refs>): MaybePromise<readonly WorkbookRunReadback[]>
  verifyChecks?(checks: readonly WorkbookCheckResult[], plan: WorkbookActionPlan<Refs>): MaybePromise<readonly WorkbookCheckResult[]>
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function runError(code: WorkbookRunErrorCode, message: string): WorkbookRunError {
  return {
    code,
    message,
  }
}

function failedRun(args: {
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

function changedAfterApply(
  plan: WorkbookActionPlan,
  result: Pick<WorkbookRunApplyResult, 'appliedOps' | 'undo'>,
): readonly WorkbookChangeSummary[] {
  if (result.appliedOps !== undefined) {
    return result.appliedOps.length > 0 || result.undo !== undefined ? plan.changed : []
  }
  return plan.changed
}

function changedAfterFailedApply(
  plan: WorkbookActionPlan,
  result: Pick<WorkbookRunApplyResult, 'appliedOps' | 'undo'>,
): readonly WorkbookChangeSummary[] {
  return result.undo !== undefined || (result.appliedOps !== undefined && result.appliedOps.length > 0) ? plan.changed : []
}

function failedApplyResult(plan: WorkbookActionPlan, result: WorkbookRunApplyResult): WorkbookRunResult {
  const errors =
    result.errors !== undefined && result.errors.length > 0
      ? result.errors
      : [runError('apply_failed', `Workbook action ${plan.modelName}.${plan.actionName} failed to apply`)]
  const apply = describeApply(result)

  return failedRun({
    errors,
    apply,
    changed: changedAfterFailedApply(plan, result),
    checks: plan.checks,
    ...(result.undo !== undefined ? { undo: result.undo } : {}),
  })
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
    ...(check.proof !== undefined ? { proof: normalizeWorkbookActionInput(check.proof) } : {}),
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

function isWorkbookRunError(value: unknown): value is WorkbookRunError {
  return isRecord(value) && isWorkbookRunErrorCode(value['code']) && typeof value['message'] === 'string'
}

function isWorkbookUndoRef(value: unknown): value is WorkbookUndoRef {
  return isRecord(value) && typeof value['id'] === 'string'
}

function isWorkbookCheckResult(value: unknown): value is WorkbookCheckResult {
  return isRecord(value) && isCheckStatus(value['status']) && typeof value['kind'] === 'string' && typeof value['message'] === 'string'
}

type ApplyValidation =
  | {
      readonly status: 'valid'
      readonly result: WorkbookRunApplyResult
    }
  | {
      readonly status: 'invalid'
      readonly result: WorkbookRunResult
    }

function validateApplyResult(plan: WorkbookActionPlan, value: unknown): ApplyValidation {
  const rejected = (message: string): ApplyValidation => ({
    status: 'invalid',
    result: failedRun({
      errors: [runError('runtime_rejected', message)],
      checks: plan.checks,
    }),
  })

  if (!isRecord(value)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned an invalid apply result`)
  }

  const status = value['status']
  if (status !== 'applied' && status !== 'failed') {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned an invalid apply status`)
  }

  const rawPreviewOps = value['previewOps']
  if (rawPreviewOps !== undefined && (!Array.isArray(rawPreviewOps) || !rawPreviewOps.every(isWorkbookOp))) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid preview ops`)
  }
  const rawAppliedOps = value['appliedOps']
  if (rawAppliedOps !== undefined && (!Array.isArray(rawAppliedOps) || !rawAppliedOps.every(isWorkbookOp))) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid applied ops`)
  }

  const rawErrors = value['errors']
  if (rawErrors !== undefined && (!Array.isArray(rawErrors) || !rawErrors.every(isWorkbookRunError))) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid apply errors`)
  }
  const errors = rawErrors as readonly WorkbookRunError[] | undefined

  if (status === 'applied' && errors !== undefined && errors.length > 0) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned applied with errors`)
  }

  const undo = value['undo']
  if (undo !== undefined && !isWorkbookUndoRef(undo)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid undo metadata`)
  }

  let proof: WorkbookActionInput | undefined
  try {
    proof = value['proof'] === undefined ? undefined : normalizeWorkbookActionInput(value['proof'])
  } catch (error) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid apply proof: ${errorMessage(error)}`)
  }

  return {
    status: 'valid',
    result: {
      status,
      ...(rawPreviewOps !== undefined ? { previewOps: cloneOps(rawPreviewOps as readonly EngineOp[]) } : {}),
      ...(rawAppliedOps !== undefined ? { appliedOps: cloneOps(rawAppliedOps as readonly EngineOp[]) } : {}),
      ...(proof !== undefined ? { proof } : {}),
      ...(errors !== undefined ? { errors } : {}),
      ...(undo !== undefined ? { undo } : {}),
    },
  }
}

function cloneOps(ops: readonly EngineOp[]): readonly EngineOp[] {
  return ops.map((op) => structuredClone(op))
}

function opsMatch(left: readonly EngineOp[], right: readonly EngineOp[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  return left.every((op, index) => {
    const other = right[index]
    return other !== undefined && canonicalJson(op) === canonicalJson(other)
  })
}

function describeApply(result: WorkbookRunApplyResult): WorkbookRunApplySummary {
  const matched = result.previewOps === undefined || result.appliedOps === undefined ? null : opsMatch(result.previewOps, result.appliedOps)
  return {
    matched,
    ...(result.previewOps !== undefined ? { previewOps: cloneOps(result.previewOps) } : {}),
    ...(result.appliedOps !== undefined ? { appliedOps: cloneOps(result.appliedOps) } : {}),
    ...(result.proof !== undefined ? { proof: normalizeWorkbookActionInput(result.proof) } : {}),
  }
}

function applyUnverified(apply: WorkbookRunApplySummary): readonly WorkbookRunUnverified[] {
  if (apply.matched !== null) {
    return []
  }
  return [
    {
      kind: 'apply',
      message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
    },
  ]
}

function unverifiedProperty(unverified: readonly WorkbookRunUnverified[]): { readonly unverified: readonly WorkbookRunUnverified[] } | {} {
  return unverified.length === 0 ? {} : { unverified }
}

function applyProofErrors(apply: WorkbookRunApplySummary, options: WorkbookRunOptions): readonly WorkbookRunError[] {
  if (apply.matched === false) {
    return [runError('apply_mismatch', 'Adapter applied ops do not match its preview ops')]
  }
  if (options.requireApplyProof === true && apply.matched === null) {
    return [runError('apply_not_verified', 'Adapter did not return both previewOps and appliedOps')]
  }
  return []
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

function changedAfterOptionalApply(plan: WorkbookActionPlan, result: WorkbookRunApplyResult | undefined): readonly WorkbookChangeSummary[] {
  return result === undefined ? [] : changedAfterApply(plan, result)
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
    try {
      verifiedChecks.push(cloneCheck(actual))
    } catch (error) {
      return {
        status: 'invalid',
        error: runError(
          'invalid_check_verification',
          `Check verifier returned invalid proof at index ${String(index)}: ${errorMessage(error)}`,
        ),
      }
    }
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

async function runLiveWorkbookPlan<Refs>(
  plan: WorkbookActionPlan<Refs>,
  adapter: WorkbookRunAdapter<Refs>,
  options: WorkbookRunOptions = {},
): Promise<WorkbookRunResult> {
  const invalidPlan = failedFromPlanIssues(plan)
  if (invalidPlan !== null) {
    return invalidPlan
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
  if (needsApply(adapterCheck.requiredCapabilities)) {
    let applyResult: unknown
    try {
      applyResult = await adapter.apply?.(plan)
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
    unverified = applyUnverified(apply)
    const applyErrors = applyProofErrors(apply, options)
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
    if (adapter.read === undefined) {
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
      readbacks = await adapter.read(targets, plan)
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
export function runWorkbookPlan<Refs>(
  plan: WorkbookExecutablePlan<Refs>,
  adapter: WorkbookRunAdapter,
  options: WorkbookRunOptions = {},
): Promise<WorkbookRunResult> {
  if (isHydratedPlan(plan)) {
    return runLiveWorkbookPlan<unknown>(plan, adapter, options)
  }
  return runLiveWorkbookPlan<unknown>(hydratePlanData(plan), adapter, options)
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
    return failedRun({
      errors: result.errors,
      checks: result.checks,
    })
  }
  return runLiveWorkbookPlan(result.plan, adapter, options)
}
