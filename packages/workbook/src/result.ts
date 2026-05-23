import type { LiteralInput } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'
import type { WorkbookActionInput } from './input.js'
import type { EngineOp } from './ops.js'

export type WorkbookCheckStatus = 'planned' | 'passed' | 'failed'

export type WorkbookCheckExpectation =
  | {
      readonly kind: 'valueEquals'
      readonly value: LiteralInput
    }
  | {
      readonly kind: 'formulaEquals'
      readonly formula: string
      readonly inputs: readonly WorkbookRef[]
    }

export interface WorkbookCheckResult {
  readonly status: WorkbookCheckStatus
  readonly kind: string
  readonly target?: WorkbookRef
  readonly refs?: readonly WorkbookRef[]
  readonly message: string
  readonly expectation?: WorkbookCheckExpectation
  readonly proof?: WorkbookActionInput
}

export interface WorkbookChangeSummary {
  readonly kind: string
  readonly target?: WorkbookRef
  readonly message: string
}

export interface WorkbookUndoRef {
  readonly id: string
  readonly ops?: readonly EngineOp[]
}

export interface WorkbookRunApplySummary {
  readonly matched: boolean | null
  readonly previewOps?: readonly EngineOp[]
  readonly appliedOps?: readonly EngineOp[]
  readonly proof?: WorkbookActionInput
}

export type WorkbookRunUnverifiedKind = 'apply'

export interface WorkbookRunUnverified {
  readonly kind: WorkbookRunUnverifiedKind
  readonly message: string
}

export type WorkbookRunErrorCode =
  | 'action_not_found'
  | 'invalid_action_input'
  | 'find_failed'
  | 'checks_failed'
  | 'action_failed'
  | 'ref_not_in_refs'
  | 'duplicate_ref'
  | 'command_target_not_resolved'
  | 'formula_input_not_resolved'
  | 'invalid_formula'
  | 'change_target_not_resolved'
  | 'check_status_not_planned'
  | 'check_target_not_resolved'
  | 'check_ref_not_resolved'
  | 'check_expectation_input_not_resolved'
  | 'invalid_check_expectation_formula'
  | 'invalid_workbook_op'
  | 'op_target_mismatch'
  | 'missing_concrete_op'
  | 'missing_workbook_op'
  | 'apply_not_verified'
  | 'apply_mismatch'
  | 'apply_failed'
  | 'readback_failed'
  | 'readback_missing'
  | 'readback_unexpected'
  | 'value_mismatch'
  | 'formula_mismatch'
  | 'invalid_check_verification'
  | 'check_verification_failed'
  | 'check_failed'
  | 'check_not_verified'
  | 'runtime_rejected'

export const workbookRunErrorCodes = Object.freeze([
  'action_not_found',
  'invalid_action_input',
  'find_failed',
  'checks_failed',
  'action_failed',
  'ref_not_in_refs',
  'duplicate_ref',
  'command_target_not_resolved',
  'formula_input_not_resolved',
  'invalid_formula',
  'change_target_not_resolved',
  'check_status_not_planned',
  'check_target_not_resolved',
  'check_ref_not_resolved',
  'check_expectation_input_not_resolved',
  'invalid_check_expectation_formula',
  'invalid_workbook_op',
  'op_target_mismatch',
  'missing_concrete_op',
  'missing_workbook_op',
  'apply_not_verified',
  'apply_mismatch',
  'apply_failed',
  'readback_failed',
  'readback_missing',
  'readback_unexpected',
  'value_mismatch',
  'formula_mismatch',
  'invalid_check_verification',
  'check_verification_failed',
  'check_failed',
  'check_not_verified',
  'runtime_rejected',
] satisfies readonly WorkbookRunErrorCode[])

export function isWorkbookRunErrorCode(value: unknown): value is WorkbookRunErrorCode {
  return typeof value === 'string' && workbookRunErrorCodes.some((code) => code === value)
}

export interface WorkbookRunError {
  readonly code: WorkbookRunErrorCode
  readonly message: string
}

export type WorkbookRunResult =
  | {
      readonly status: 'done'
      readonly apply?: WorkbookRunApplySummary
      readonly changed: readonly WorkbookChangeSummary[]
      readonly checks: readonly WorkbookCheckResult[]
      readonly undo?: WorkbookUndoRef
      readonly unverified?: readonly WorkbookRunUnverified[]
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunError[]
      readonly apply?: WorkbookRunApplySummary
      readonly changed: readonly WorkbookChangeSummary[]
      readonly checks: readonly WorkbookCheckResult[]
      readonly undo?: WorkbookUndoRef
      readonly unverified?: readonly WorkbookRunUnverified[]
    }
