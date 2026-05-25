import { isWorkbookOp } from './guards.js'
import { normalizeWorkbookActionInput, type WorkbookActionInput } from './input.js'
import type { WorkbookActionPlan } from './model.js'
import type { EngineOp } from './ops.js'
import { workbookPlanId } from './plan-data.js'
import { cloneWorkbookRunApplyCommandReceipts, cloneWorkbookRunApplyCommandReceiptsForSummary } from './run-command-receipts.js'
import { cloneData, arrayDataValues, canonicalJson, isRecord, ownValue } from './run-data.js'
import { errorMessage, failedRun, runError } from './run-failure.js'
import type {
  WorkbookChangeSummary,
  WorkbookRunApplyCommandReceipt,
  WorkbookRunApplySummary,
  WorkbookRunError,
  WorkbookRunResult,
  WorkbookRunUnverified,
  WorkbookUndoRef,
} from './result.js'
import { isWorkbookRunErrorCode } from './result.js'
import type { WorkbookRunApplyResult } from './run.js'

type ApplyValidation =
  | {
      readonly status: 'valid'
      readonly result: WorkbookRunApplyResult
    }
  | {
      readonly status: 'invalid'
      readonly result: WorkbookRunResult
    }

export function changedAfterApply(
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

export function changedAfterOptionalApply(
  plan: WorkbookActionPlan,
  result: WorkbookRunApplyResult | undefined,
): readonly WorkbookChangeSummary[] {
  return result === undefined ? [] : changedAfterApply(plan, result)
}

export function failedApplyResult(plan: WorkbookActionPlan, result: WorkbookRunApplyResult): WorkbookRunResult {
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

export function validateApplyResult(plan: WorkbookActionPlan, value: unknown): ApplyValidation {
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

export function describeApply(result: WorkbookRunApplyResult): WorkbookRunApplySummary {
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

export function applyUnverified(plan: WorkbookActionPlan, apply: WorkbookRunApplySummary): readonly WorkbookRunUnverified[] {
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

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
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
