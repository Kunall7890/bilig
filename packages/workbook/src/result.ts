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
      readonly kind: 'valuesEqual'
      readonly values: readonly (readonly LiteralInput[])[]
    }
  | {
      readonly kind: 'formulaEquals'
      readonly formula: string
      readonly inputs: readonly WorkbookRef[]
    }
  | {
      readonly kind: 'formulasEqual'
      readonly formulas: readonly (readonly (string | null)[])[]
    }

export interface WorkbookCheckResult {
  readonly status: WorkbookCheckStatus
  readonly kind: string
  readonly target?: WorkbookRef
  readonly refs?: readonly WorkbookRef[]
  readonly message: string
  readonly expectation?: WorkbookCheckExpectation
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

export interface WorkbookAppliedSummary {
  readonly opCount: number
  readonly ops?: readonly EngineOp[]
}

export type WorkbookRunErrorCode =
  | 'action_not_found'
  | 'invalid_action_input'
  | 'find_failed'
  | 'checks_failed'
  | 'action_failed'
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
  | 'apply_failed'
  | 'readback_failed'
  | 'readback_missing'
  | 'duplicate_readback'
  | 'value_mismatch'
  | 'values_mismatch'
  | 'formula_mismatch'
  | 'formulas_mismatch'
  | 'invalid_check_verification'
  | 'check_verification_failed'
  | 'check_failed'
  | 'check_not_verified'
  | 'invalid_runtime_result'
  | 'runtime_rejected'

export const workbookRunErrorCodes = Object.freeze([
  'action_not_found',
  'invalid_action_input',
  'find_failed',
  'checks_failed',
  'action_failed',
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
  'apply_failed',
  'readback_failed',
  'readback_missing',
  'duplicate_readback',
  'value_mismatch',
  'values_mismatch',
  'formula_mismatch',
  'formulas_mismatch',
  'invalid_check_verification',
  'check_verification_failed',
  'check_failed',
  'check_not_verified',
  'invalid_runtime_result',
  'runtime_rejected',
] satisfies readonly WorkbookRunErrorCode[])

export function isWorkbookRunErrorCode(value: unknown): value is WorkbookRunErrorCode {
  return typeof value === 'string' && workbookRunErrorCodes.some((code) => code === value)
}

export interface WorkbookRunError {
  readonly code: WorkbookRunErrorCode
  readonly message: string
  readonly path?: string
  readonly target?: WorkbookRef
  readonly check?: WorkbookCheckResult
  readonly expected?: WorkbookActionInput
  readonly actual?: WorkbookActionInput
}

export type WorkbookRunResult =
  | {
      readonly status: 'done'
      readonly changed: readonly WorkbookChangeSummary[]
      readonly checks: readonly WorkbookCheckResult[]
      readonly undo?: WorkbookUndoRef
      readonly applied?: WorkbookAppliedSummary
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunError[]
      readonly checks: readonly WorkbookCheckResult[]
      readonly undo?: WorkbookUndoRef
    }
