import { isLiteralInput, type LiteralInput } from '@bilig/protocol'
import { isWorkbookRef, type WorkbookRef } from './find.js'
import { isWorkbookOp } from './guards.js'
import type { WorkbookActionInput } from './input.js'
import { planWorkbookAction, type WorkbookActionMap, type WorkbookActionPlan, type WorkbookModel } from './model.js'
import { verifyWorkbookReadbacks, type WorkbookCellReadback, type WorkbookReadbackIssue, type WorkbookRunReadback } from './readback.js'
import type { WorkbookRuntimePreview, WorkbookRuntimeRequirement } from './requirements.js'
import {
  isWorkbookRunErrorCode,
  type WorkbookAppliedSummary,
  type WorkbookCheckProof,
  type WorkbookCheckResult,
  type WorkbookRunError,
  type WorkbookRunErrorCode,
  type WorkbookRunResult,
  type WorkbookUndoRef,
} from './result.js'
import { verifyPlan } from './verify.js'

type MaybePromise<T> = T | Promise<T>

export interface WorkbookRunApplyResult {
  readonly status: 'applied' | 'failed'
  readonly errors?: readonly WorkbookRunError[]
  readonly undo?: WorkbookUndoRef
}

export interface WorkbookRunAdapter<Refs = unknown> {
  readonly preview?: (plan: WorkbookActionPlan<Refs>) => MaybePromise<WorkbookRuntimePreview>
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

function runError(code: WorkbookRunErrorCode, message: string): WorkbookRunError {
  return {
    code,
    message,
  }
}

function runtimeResultError(message: string, path = 'runtime'): WorkbookRunError {
  return {
    code: 'invalid_runtime_result',
    message,
    path,
  }
}

function planIssueError(issue: ReturnType<typeof verifyPlan>['issues'][number]): WorkbookRunError {
  return {
    code: issue.code,
    message: issue.message,
    path: issue.path,
  }
}

function readbackIssueError(issue: WorkbookReadbackIssue): WorkbookRunError {
  return {
    code: issue.code,
    message: issue.message,
    ...(issue.path !== undefined ? { path: issue.path } : {}),
    ...(issue.target !== undefined ? { target: issue.target } : {}),
    check: issue.check,
    ...(issue.expected !== undefined ? { expected: issue.expected } : {}),
    ...(issue.actual !== undefined ? { actual: issue.actual } : {}),
  }
}

function appliedSummary(preview: WorkbookRuntimePreview | undefined): WorkbookAppliedSummary | undefined {
  if (preview === undefined) {
    return undefined
  }
  return {
    opCount: preview.materializedOps.length,
    ops: preview.materializedOps,
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
    errors: verification.issues.map(planIssueError),
    checks: plan.checks,
  }
}

function failedApplyResult(plan: WorkbookActionPlan, result: WorkbookRunApplyResult): WorkbookRunResult {
  return {
    status: 'failed',
    errors:
      result.errors !== undefined && result.errors.length > 0
        ? result.errors
        : [runError('apply_failed', `Workbook action ${plan.modelName}.${plan.actionName} failed to apply`)],
    checks: plan.checks,
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
    ...(check.proof !== undefined ? { proof: check.proof } : {}),
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
    ...(check.proof !== undefined ? { proof: check.proof } : {}),
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

function isStringOrNull(value: unknown): value is string | null {
  return typeof value === 'string' || value === null
}

function isRectangularMatrix(value: readonly unknown[][]): boolean {
  const width = value[0]?.length
  return value.every((row) => row.length === width)
}

function isLiteralMatrix(value: unknown): value is readonly (readonly LiteralInput[])[] {
  return (
    Array.isArray(value) &&
    value.every((row) => Array.isArray(row) && row.every((entry) => isLiteralInput(entry))) &&
    isRectangularMatrix(value)
  )
}

function isFormulaMatrix(value: unknown): value is readonly (readonly (string | null)[])[] {
  return Array.isArray(value) && value.every((row) => Array.isArray(row) && row.every(isStringOrNull)) && isRectangularMatrix(value)
}

function isWorkbookCheckProof(value: unknown): value is WorkbookCheckProof {
  if (!isRecord(value)) {
    return false
  }
  switch (value['kind']) {
    case 'value':
      return isLiteralInput(value['value'])
    case 'values':
      return isLiteralMatrix(value['values'])
    case 'formula':
      return isStringOrNull(value['formula'])
    case 'formulas':
      return isFormulaMatrix(value['formulas'])
    default:
      return false
  }
}

function normalizedCellReadback(value: unknown): WorkbookCellReadback | null {
  if (!isRecord(value) || typeof value['sheetName'] !== 'string' || typeof value['address'] !== 'string') {
    return null
  }
  const cell: { sheetName: string; address: string; value?: LiteralInput; formula?: string | null } = {
    sheetName: value['sheetName'],
    address: value['address'],
  }
  if (value['value'] !== undefined) {
    if (!isLiteralInput(value['value'])) {
      return null
    }
    cell.value = value['value']
  }
  if (value['formula'] !== undefined) {
    if (!isStringOrNull(value['formula'])) {
      return null
    }
    cell.formula = value['formula']
  }
  return cell
}

function isPreviewRequirement(value: unknown): value is WorkbookRuntimeRequirement {
  return (
    isRecord(value) &&
    (value['kind'] === 'apply' || value['kind'] === 'read' || value['kind'] === 'verify') &&
    typeof value['capability'] === 'string' &&
    typeof value['path'] === 'string' &&
    typeof value['message'] === 'string'
  )
}

function isRunError(value: unknown): value is WorkbookRunError {
  return isRecord(value) && isWorkbookRunErrorCode(value['code']) && typeof value['message'] === 'string'
}

function validatePreviewResult(value: unknown, plan: WorkbookActionPlan): WorkbookRuntimePreview | WorkbookRunError {
  if (!isRecord(value)) {
    return runtimeResultError('Runtime preview must return an object', 'preview')
  }
  if (value['modelName'] !== plan.modelName) {
    return runtimeResultError(`Runtime preview returned modelName ${String(value['modelName'])} for ${plan.modelName}`, 'preview.modelName')
  }
  if (value['actionName'] !== plan.actionName) {
    return runtimeResultError(
      `Runtime preview returned actionName ${String(value['actionName'])} for ${plan.actionName}`,
      'preview.actionName',
    )
  }
  if (!Array.isArray(value['requirements'])) {
    return runtimeResultError('Runtime preview requirements must be an array', 'preview.requirements')
  }
  const requirements = value['requirements']
  if (!requirements.every(isPreviewRequirement)) {
    return runtimeResultError('Runtime preview requirements must be runtime requirement objects', 'preview.requirements')
  }
  const materializedOps = value['materializedOps']
  if (!Array.isArray(materializedOps) || !materializedOps.every(isWorkbookOp)) {
    return runtimeResultError('Runtime preview materializedOps must be a WorkbookOp array', 'preview.materializedOps')
  }
  return {
    modelName: plan.modelName,
    actionName: plan.actionName,
    requirements,
    materializedOps,
  }
}

function validateApplyResult(value: unknown): WorkbookRunApplyResult | WorkbookRunError {
  if (!isRecord(value)) {
    return runtimeResultError('Runtime apply must return an object', 'apply')
  }
  if (value['status'] !== 'applied' && value['status'] !== 'failed') {
    return runtimeResultError('Runtime apply status must be applied or failed', 'apply.status')
  }
  const errors = value['errors']
  if (errors !== undefined && (!Array.isArray(errors) || !errors.every(isRunError))) {
    return runtimeResultError('Runtime apply errors must be WorkbookRunError objects', 'apply.errors')
  }
  const undo = value['undo']
  let undoRef: WorkbookUndoRef | undefined
  if (undo !== undefined) {
    if (!isRecord(undo) || typeof undo['id'] !== 'string') {
      return runtimeResultError('Runtime apply undo must include an id', 'apply.undo')
    }
    if (undo['ops'] !== undefined && (!Array.isArray(undo['ops']) || !undo['ops'].every(isWorkbookOp))) {
      return runtimeResultError('Runtime apply undo ops must be a WorkbookOp array', 'apply.undo.ops')
    }
    undoRef = {
      id: undo['id'],
      ...(undo['ops'] !== undefined ? { ops: undo['ops'] } : {}),
    }
  }
  return {
    status: value['status'],
    ...(errors !== undefined ? { errors } : {}),
    ...(undoRef !== undefined ? { undo: undoRef } : {}),
  }
}

function validateReadbacks(value: unknown): readonly WorkbookRunReadback[] | WorkbookRunError {
  if (!Array.isArray(value)) {
    return runtimeResultError('Runtime read must return a readback array', 'read')
  }
  const readbacks: WorkbookRunReadback[] = []
  for (let index = 0; index < value.length; index += 1) {
    const entry = value[index]
    const path = `read[${index.toString()}]`
    if (!isRecord(entry)) {
      return runtimeResultError('Runtime readback must be an object', path)
    }
    if (!isWorkbookRef(entry['target'])) {
      return runtimeResultError('Runtime readback target must be a WorkbookRef', `${path}.target`)
    }
    const readback: {
      target: WorkbookRef
      value?: LiteralInput
      formula?: string | null
      values?: readonly (readonly LiteralInput[])[]
      formulas?: readonly (readonly (string | null)[])[]
      cells?: readonly WorkbookCellReadback[]
    } = {
      target: entry['target'],
    }
    if (entry['value'] !== undefined && !isLiteralInput(entry['value'])) {
      return runtimeResultError('Runtime readback value must be a literal input', `${path}.value`)
    }
    if (entry['value'] !== undefined) {
      readback.value = entry['value']
    }
    if (entry['formula'] !== undefined && !isStringOrNull(entry['formula'])) {
      return runtimeResultError('Runtime readback formula must be a string or null', `${path}.formula`)
    }
    if (entry['formula'] !== undefined) {
      readback.formula = entry['formula']
    }
    if (entry['values'] !== undefined && !isLiteralMatrix(entry['values'])) {
      return runtimeResultError('Runtime readback values must be an array of literal rows', `${path}.values`)
    }
    if (entry['values'] !== undefined) {
      readback.values = entry['values']
    }
    if (entry['formulas'] !== undefined && !isFormulaMatrix(entry['formulas'])) {
      return runtimeResultError('Runtime readback formulas must be an array of string/null rows', `${path}.formulas`)
    }
    if (entry['formulas'] !== undefined) {
      readback.formulas = entry['formulas']
    }
    if (entry['cells'] !== undefined) {
      if (!Array.isArray(entry['cells'])) {
        return runtimeResultError('Runtime readback cells must be cell readback objects', `${path}.cells`)
      }
      const cells: WorkbookCellReadback[] = []
      for (let cellIndex = 0; cellIndex < entry['cells'].length; cellIndex += 1) {
        const cell = normalizedCellReadback(entry['cells'][cellIndex])
        if (cell === null) {
          return runtimeResultError('Runtime readback cells must be cell readback objects', `${path}.cells`)
        }
        cells.push(cell)
      }
      readback.cells = cells
    }
    readbacks.push(readback)
  }
  return readbacks
}

function isWorkbookCheckResult(value: unknown): value is WorkbookCheckResult {
  return (
    isRecord(value) &&
    isCheckStatus(value['status']) &&
    typeof value['kind'] === 'string' &&
    typeof value['message'] === 'string' &&
    (value['proof'] === undefined || isWorkbookCheckProof(value['proof']))
  )
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

function failedAfterApply(
  applyResult: WorkbookRunApplyResult,
  errors: readonly WorkbookRunError[],
  checks: readonly WorkbookCheckResult[],
): WorkbookRunResult {
  return {
    status: 'failed',
    errors,
    checks,
    ...(applyResult.undo !== undefined ? { undo: applyResult.undo } : {}),
  }
}

export async function runWorkbookPlan<Refs>(plan: WorkbookActionPlan<Refs>, adapter: WorkbookRunAdapter<Refs>): Promise<WorkbookRunResult> {
  const invalidPlan = failedFromPlanIssues(plan)
  if (invalidPlan !== null) {
    return invalidPlan
  }

  let preview: WorkbookRuntimePreview | undefined
  if (adapter.preview !== undefined) {
    try {
      const previewResult = validatePreviewResult(await adapter.preview(plan), plan)
      if ('code' in previewResult) {
        return {
          status: 'failed',
          errors: [previewResult],
          checks: plan.checks,
        }
      }
      preview = previewResult
    } catch (error) {
      return {
        status: 'failed',
        errors: [runError('runtime_rejected', errorMessage(error))],
        checks: plan.checks,
      }
    }
  }

  let applyResult: WorkbookRunApplyResult
  try {
    const rawApplyResult = validateApplyResult(await adapter.apply(plan))
    if ('code' in rawApplyResult) {
      return {
        status: 'failed',
        errors: [rawApplyResult],
        checks: plan.checks,
      }
    }
    applyResult = rawApplyResult
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

  let checks = plan.checks
  const targets = readbackTargets(checks)
  if (targets.length > 0) {
    if (adapter.read === undefined) {
      const readbackVerification = verifyWorkbookReadbacks(checks, [])
      return failedAfterApply(applyResult, readbackVerification.issues.map(readbackIssueError), readbackVerification.checks)
    }

    let readbacks: readonly WorkbookRunReadback[]
    try {
      const rawReadbacks = validateReadbacks(await adapter.read(targets, plan))
      if ('code' in rawReadbacks) {
        return failedAfterApply(applyResult, [rawReadbacks], checks)
      }
      readbacks = rawReadbacks
    } catch (error) {
      return failedAfterApply(applyResult, [runError('readback_failed', errorMessage(error))], checks)
    }

    const readbackVerification = verifyWorkbookReadbacks(checks, readbacks)
    checks = readbackVerification.checks
    if (readbackVerification.status === 'failed') {
      return failedAfterApply(applyResult, readbackVerification.issues.map(readbackIssueError), checks)
    }
  }

  const checkVerification = await verifyChecksWithAdapter(checks, plan, adapter)
  checks = checkVerification.checks
  if (checkVerification.errors.length > 0) {
    return failedAfterApply(applyResult, checkVerification.errors, checks)
  }

  const unverifiedErrors = unverifiedCheckErrors(checks)
  if (unverifiedErrors.length > 0) {
    return failedAfterApply(applyResult, unverifiedErrors, checks)
  }

  const applied = appliedSummary(preview)
  return {
    status: 'done',
    changed: plan.changed,
    checks,
    ...(applyResult.undo !== undefined ? { undo: applyResult.undo } : {}),
    ...(applied !== undefined ? { applied } : {}),
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
