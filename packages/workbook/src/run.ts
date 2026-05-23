import type { WorkbookRef } from './find.js'
import { isWorkbookOp } from './guards.js'
import { normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import { planWorkbookAction, type WorkbookActionMap, type WorkbookActionPlan, type WorkbookModel } from './model.js'
import type { EngineOp } from './ops.js'
import {
  checkPlanData,
  hydratePlanData,
  isHydratedPlan,
  workbookPlanId,
  type WorkbookExecutablePlan,
  type WorkbookPlanId,
  type WorkbookPlanDataIssue,
  type WorkbookPlanDataRefs,
} from './plan-data.js'
import { verifyWorkbookReadbacks, type WorkbookRunReadback } from './readback.js'
import { checkRuntimeAdapter, type WorkbookRuntimeCapability } from './requirements.js'
import { cloneWorkbookRunApplyCommandReceipts, cloneWorkbookRunApplyCommandReceiptsForSummary } from './run-command-receipts.js'
import {
  isWorkbookRunErrorCode,
  type WorkbookChangeSummary,
  type WorkbookRunApplySummary,
  type WorkbookRunApplyCommandReceipt,
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
  readonly requireApplyProof?: boolean
  readonly requirePlanId?: boolean
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

function planDataRunError(issue: WorkbookPlanDataIssue): WorkbookRunError {
  return {
    code: 'invalid_plan_data',
    message: issue.message,
    path: issue.path,
    issueCode: issue.code,
  }
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
    const descriptors = Object.getOwnPropertyDescriptors(value)
    return Array.from({ length: value.length }, (_, index) => {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        return undefined
      }
      if (!('value' in descriptor)) {
        throw new Error('Accessor values cannot be canonicalized')
      }
      return canonicalValue(descriptor.value)
    })
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(Object.getOwnPropertyDescriptors(value))
        .filter(([, descriptor]) => descriptor.enumerable)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, descriptor]) => {
          if (!('value' in descriptor)) {
            throw new Error('Accessor values cannot be canonicalized')
          }
          return [key, canonicalValue(descriptor.value)]
        }),
    )
  }
  return value
}

function ownCheckValue<Key extends keyof WorkbookCheckResult>(check: WorkbookCheckResult, key: Key): WorkbookCheckResult[Key] | undefined {
  const descriptor = Object.getOwnPropertyDescriptor(check, key)
  if (descriptor === undefined) {
    return undefined
  }
  if (!('value' in descriptor)) {
    throw new Error(`Workbook check result ${String(key)} must be a data property`)
  }
  return descriptor.value
}

function checkContract(check: WorkbookCheckResult): Omit<WorkbookCheckResult, 'status'> {
  const kind = ownCheckValue(check, 'kind')
  const target = ownCheckValue(check, 'target')
  const refs = ownCheckValue(check, 'refs')
  const message = ownCheckValue(check, 'message')
  const expectation = ownCheckValue(check, 'expectation')
  if (kind === undefined || message === undefined) {
    throw new Error('invalid check contract')
  }

  return {
    kind,
    ...(target !== undefined ? { target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message,
    ...(expectation !== undefined ? { expectation } : {}),
  }
}

function cloneCheck(check: WorkbookCheckResult): WorkbookCheckResult {
  const status = ownCheckValue(check, 'status')
  const kind = ownCheckValue(check, 'kind')
  const target = ownCheckValue(check, 'target')
  const refs = ownCheckValue(check, 'refs')
  const message = ownCheckValue(check, 'message')
  const expectation = ownCheckValue(check, 'expectation')
  const proof = ownCheckValue(check, 'proof')
  if (status === undefined || kind === undefined || message === undefined) {
    throw new Error('invalid check result')
  }

  return {
    status,
    kind,
    ...(target !== undefined ? { target } : {}),
    ...(refs !== undefined ? { refs } : {}),
    message,
    ...(expectation !== undefined ? { expectation } : {}),
    ...(proof !== undefined ? { proof: normalizeWorkbookActionInput(proof) } : {}),
  }
}

function checkContractMatches(expectedContract: string, actual: WorkbookCheckResult): boolean {
  try {
    return expectedContract === canonicalJson(checkContract(actual))
  } catch {
    return false
  }
}

function isCheckStatus(value: unknown): value is WorkbookCheckResult['status'] {
  return value === 'planned' || value === 'passed' || value === 'failed'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function ownValue(value: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(value, key)?.value
}

function arrayDataValues<T>(value: unknown, guard: (entry: unknown) => entry is T): readonly T[] | null {
  if (!Array.isArray(value)) {
    return null
  }
  if (firstAccessorPath(value, 'array') !== null) {
    return null
  }

  const entries: T[] = []
  for (let index = 0; index < value.length; index += 1) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index))
    if (descriptor === undefined || !descriptor.enumerable || !('value' in descriptor) || !guard(descriptor.value)) {
      return null
    }
    entries.push(descriptor.value)
  }
  return entries
}

function isWorkbookOpArray(value: unknown): value is readonly EngineOp[] {
  return arrayDataValues(value, isWorkbookOp) !== null
}

function isWorkbookRunError(value: unknown): value is WorkbookRunError {
  if (!isRecord(value)) {
    return false
  }
  const code = ownValue(value, 'code')
  const message = ownValue(value, 'message')
  const path = ownValue(value, 'path')
  const issueCode = ownValue(value, 'issueCode')

  return (
    isWorkbookRunErrorCode(code) &&
    typeof message === 'string' &&
    (path === undefined || typeof path === 'string') &&
    (issueCode === undefined || typeof issueCode === 'string')
  )
}

function isWorkbookRunErrorArray(value: unknown): value is readonly WorkbookRunError[] {
  return arrayDataValues(value, isWorkbookRunError) !== null
}

function cloneRunError(error: WorkbookRunError): WorkbookRunError {
  const code = ownValue(error, 'code')
  const message = ownValue(error, 'message')
  const path = ownValue(error, 'path')
  const issueCode = ownValue(error, 'issueCode')
  if (!isWorkbookRunErrorCode(code) || typeof message !== 'string') {
    throw new Error('invalid run error')
  }

  return {
    code,
    message,
    ...(typeof path === 'string' ? { path } : {}),
    ...(typeof issueCode === 'string' ? { issueCode } : {}),
  }
}

function isWorkbookUndoRef(value: unknown): value is WorkbookUndoRef {
  if (!isRecord(value)) {
    return false
  }
  const id = ownValue(value, 'id')
  const ops = ownValue(value, 'ops')
  return typeof id === 'string' && (ops === undefined || isWorkbookOpArray(ops))
}

function cloneUndoRef(undo: WorkbookUndoRef): WorkbookUndoRef {
  const id = ownValue(undo, 'id')
  const ops = ownValue(undo, 'ops')
  if (typeof id !== 'string') {
    throw new Error('invalid undo metadata')
  }
  return {
    id,
    ...(isWorkbookOpArray(ops) ? { ops: cloneOps(ops) } : {}),
  }
}

function isWorkbookCheckResult(value: unknown): value is WorkbookCheckResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    isCheckStatus(ownValue(value, 'status')) &&
    typeof ownValue(value, 'kind') === 'string' &&
    typeof ownValue(value, 'message') === 'string'
  )
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

  const status = ownValue(value, 'status')
  if (status !== 'applied' && status !== 'failed') {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned an invalid apply status`)
  }

  const rawPreviewOps = ownValue(value, 'previewOps')
  if (rawPreviewOps !== undefined && !isWorkbookOpArray(rawPreviewOps)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid preview ops`)
  }
  let previewOps: readonly EngineOp[] | undefined
  try {
    previewOps = rawPreviewOps === undefined ? undefined : cloneOps(rawPreviewOps)
  } catch {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid preview ops`)
  }
  const rawAppliedOps = ownValue(value, 'appliedOps')
  if (rawAppliedOps !== undefined && !isWorkbookOpArray(rawAppliedOps)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid applied ops`)
  }
  let appliedOps: readonly EngineOp[] | undefined
  try {
    appliedOps = rawAppliedOps === undefined ? undefined : cloneOps(rawAppliedOps)
  } catch {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid applied ops`)
  }

  const rawErrors = ownValue(value, 'errors')
  if (rawErrors !== undefined && !isWorkbookRunErrorArray(rawErrors)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid apply errors`)
  }
  let errors: readonly WorkbookRunError[] | undefined
  try {
    errors = rawErrors === undefined ? undefined : cloneRunErrors(rawErrors)
  } catch {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid apply errors`)
  }

  if (status === 'applied' && errors !== undefined && errors.length > 0) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned applied with errors`)
  }

  const rawPlanId = ownValue(value, 'planId')
  if (rawPlanId !== undefined && typeof rawPlanId !== 'string') {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid plan id`)
  }
  const expectedPlanId = rawPlanId === undefined ? undefined : workbookPlanId(plan)
  if (rawPlanId !== undefined && rawPlanId !== expectedPlanId) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned a plan id that does not match the executed plan`)
  }

  const rawBaseRevision = ownValue(value, 'baseRevision')
  if (rawBaseRevision !== undefined && !isSafeNonNegativeInteger(rawBaseRevision)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid base revision`)
  }

  const rawRevision = ownValue(value, 'revision')
  if (rawRevision !== undefined && !isSafeNonNegativeInteger(rawRevision)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid revision`)
  }
  if (isSafeNonNegativeInteger(rawBaseRevision) && isSafeNonNegativeInteger(rawRevision) && rawRevision < rawBaseRevision) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned a revision before its base revision`)
  }

  const rawCommandReceipts = ownValue(value, 'commandReceipts')
  let commandReceipts: readonly WorkbookRunApplyCommandReceipt[] | undefined
  if (rawCommandReceipts !== undefined) {
    try {
      commandReceipts = cloneWorkbookRunApplyCommandReceipts(plan, rawCommandReceipts, previewOps, appliedOps)
    } catch (error) {
      return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid command receipts: ${errorMessage(error)}`)
    }
  }

  const rawUndo = ownValue(value, 'undo')
  let undo: WorkbookUndoRef | undefined
  if (rawUndo !== undefined && !isWorkbookUndoRef(rawUndo)) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid undo metadata`)
  }
  if (rawUndo !== undefined) {
    try {
      undo = cloneUndoRef(rawUndo)
    } catch {
      return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid undo metadata`)
    }
  }

  let proof: WorkbookActionInput | undefined
  const rawProof = ownValue(value, 'proof')
  try {
    proof = rawProof === undefined ? undefined : normalizeWorkbookActionInput(rawProof)
  } catch (error) {
    return rejected(`Workbook action ${plan.modelName}.${plan.actionName} returned invalid apply proof: ${errorMessage(error)}`)
  }

  return {
    status: 'valid',
    result: {
      status,
      ...(typeof rawPlanId === 'string' ? { planId: rawPlanId } : {}),
      ...(isSafeNonNegativeInteger(rawBaseRevision) ? { baseRevision: rawBaseRevision } : {}),
      ...(isSafeNonNegativeInteger(rawRevision) ? { revision: rawRevision } : {}),
      ...(previewOps !== undefined ? { previewOps } : {}),
      ...(appliedOps !== undefined ? { appliedOps } : {}),
      ...(commandReceipts !== undefined ? { commandReceipts } : {}),
      ...(proof !== undefined ? { proof } : {}),
      ...(errors !== undefined ? { errors } : {}),
      ...(undo !== undefined ? { undo } : {}),
    },
  }
}

function cloneOps(ops: readonly EngineOp[]): readonly EngineOp[] {
  const entries = arrayDataValues(ops, isWorkbookOp)
  if (entries === null) {
    throw new Error('invalid workbook op array')
  }
  return entries.map((op) => cloneOp(op))
}

function cloneRunErrors(errors: readonly WorkbookRunError[]): readonly WorkbookRunError[] {
  const entries = arrayDataValues(errors, isWorkbookRunError)
  if (entries === null) {
    throw new Error('invalid run error array')
  }
  return entries.map((error) => cloneRunError(error))
}

function cloneOp(op: EngineOp): EngineOp {
  const cloned = cloneData(op)
  if (!isWorkbookOp(cloned)) {
    throw new Error('invalid workbook op clone')
  }
  return cloned
}

function opsMatch(left: readonly EngineOp[], right: readonly EngineOp[]): boolean {
  const leftOps = arrayDataValues(left, isWorkbookOp)
  const rightOps = arrayDataValues(right, isWorkbookOp)
  if (leftOps === null || rightOps === null || leftOps.length !== rightOps.length) {
    return false
  }
  try {
    return leftOps.every((op, index) => {
      const other = rightOps[index]
      return other !== undefined && canonicalJson(op) === canonicalJson(other)
    })
  } catch {
    return false
  }
}

function cloneData(value: unknown, seen = new WeakMap<object, unknown>()): unknown {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  const existing = seen.get(value)
  if (existing !== undefined) {
    return existing
  }
  if (Array.isArray(value)) {
    const cloned: unknown[] = []
    seen.set(value, cloned)
    const descriptors = Object.getOwnPropertyDescriptors(value)
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)]
      if (descriptor === undefined) {
        continue
      }
      if (!('value' in descriptor)) {
        throw new Error('Transport data must not contain accessors')
      }
      cloned[index] = cloneData(descriptor.value, seen)
    }
    return cloned
  }
  const cloned: Record<string, unknown> = Object.create(Object.getPrototypeOf(value))
  seen.set(value, cloned)
  Object.entries(Object.getOwnPropertyDescriptors(value)).forEach(([key, descriptor]) => {
    if (!descriptor.enumerable) {
      return
    }
    if (!('value' in descriptor)) {
      throw new Error('Transport data must not contain accessors')
    }
    Object.defineProperty(cloned, key, {
      configurable: true,
      enumerable: true,
      value: cloneData(descriptor.value, seen),
      writable: true,
    })
  })
  return cloned
}

function firstAccessorPath(value: unknown, path: string, seen = new WeakSet<object>()): string | null {
  if (typeof value !== 'object' || value === null) {
    return null
  }
  if (seen.has(value)) {
    return null
  }
  seen.add(value)
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    const childPath = Array.isArray(value) && /^\d+$/.test(key) ? `${path}[${key}]` : `${path}.${key}`
    if (!('value' in descriptor)) {
      return childPath
    }
    const nestedPath = firstAccessorPath(descriptor.value, childPath, seen)
    if (nestedPath !== null) {
      return nestedPath
    }
  }
  return null
}

function describeApply(result: WorkbookRunApplyResult): WorkbookRunApplySummary {
  const matched = result.previewOps === undefined || result.appliedOps === undefined ? null : opsMatch(result.previewOps, result.appliedOps)
  return {
    matched,
    ...(result.planId !== undefined ? { planId: result.planId } : {}),
    ...(result.baseRevision !== undefined ? { baseRevision: result.baseRevision } : {}),
    ...(result.revision !== undefined ? { revision: result.revision } : {}),
    ...(result.previewOps !== undefined ? { previewOps: cloneOps(result.previewOps) } : {}),
    ...(result.appliedOps !== undefined ? { appliedOps: cloneOps(result.appliedOps) } : {}),
    ...(result.commandReceipts !== undefined
      ? { commandReceipts: cloneWorkbookRunApplyCommandReceiptsForSummary(result.commandReceipts) }
      : {}),
    ...(result.proof !== undefined ? { proof: normalizeWorkbookActionInput(result.proof) } : {}),
  }
}

function applyUnverified(plan: WorkbookActionPlan, apply: WorkbookRunApplySummary): readonly WorkbookRunUnverified[] {
  const unverified: WorkbookRunUnverified[] = []
  if (apply.matched === null) {
    unverified.push({
      kind: 'apply',
      message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
    })
  }
  if (plan.commands.length > 0 && apply.commandReceipts === undefined) {
    unverified.push({
      kind: 'apply',
      message: 'Adapter did not return commandReceipts, so planned commands are not bound to materialized ops',
    })
  }
  return unverified
}

function unverifiedProperty(unverified: readonly WorkbookRunUnverified[]): { readonly unverified: readonly WorkbookRunUnverified[] } | {} {
  return unverified.length === 0 ? {} : { unverified }
}

function applyProofErrors(
  plan: WorkbookActionPlan,
  apply: WorkbookRunApplySummary,
  options: WorkbookRunOptions,
): readonly WorkbookRunError[] {
  if (apply.matched === false) {
    return [runError('apply_mismatch', 'Adapter applied ops do not match its preview ops')]
  }
  if (options.requireApplyProof === true && apply.matched === null) {
    return [runError('apply_not_verified', 'Adapter did not return both previewOps and appliedOps')]
  }
  if (options.requireApplyProof === true && plan.commands.length > 0 && apply.commandReceipts === undefined) {
    return [runError('apply_not_verified', 'Adapter did not bind planned commands to materialized ops')]
  }
  if (options.requirePlanId === true && apply.planId === undefined) {
    return [runError('plan_not_verified', 'Adapter did not bind apply proof to a plan id')]
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
    const descriptor = Object.getOwnPropertyDescriptor(verified, String(index))
    const actual = descriptor !== undefined && descriptor.enumerable && 'value' in descriptor ? descriptor.value : undefined
    if (expectedContract === undefined || !isRecord(actual)) {
      return {
        status: 'invalid',
        error: runError('invalid_check_verification', `Check verifier returned an invalid check at index ${String(index)}`),
      }
    }
    if (!isCheckStatus(ownValue(actual, 'status'))) {
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
    unverified = applyUnverified(plan, apply)
    const applyErrors = applyProofErrors(plan, apply, options)
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
export function runWorkbookPlan(plan: unknown, adapter: WorkbookRunAdapter, options?: WorkbookRunOptions): Promise<WorkbookRunResult>
export function runWorkbookPlan(plan: unknown, adapter: WorkbookRunAdapter, options: WorkbookRunOptions = {}): Promise<WorkbookRunResult> {
  if (isHydratedPlan(plan)) {
    return runLiveWorkbookPlan<unknown>(plan, adapter, options)
  }
  const planDataCheck = checkPlanData(plan)
  if (planDataCheck.status === 'invalid') {
    return Promise.resolve(
      failedRun({
        errors: planDataCheck.issues.map(planDataRunError),
        checks: [],
      }),
    )
  }
  return runLiveWorkbookPlan<unknown>(hydratePlanData(planDataCheck.plan), adapter, options)
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
