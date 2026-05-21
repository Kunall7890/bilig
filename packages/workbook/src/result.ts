import type { WorkbookRef } from './find.js'

export type WorkbookCheckStatus = 'planned' | 'passed' | 'failed'

export interface WorkbookCheckResult {
  readonly status: WorkbookCheckStatus
  readonly kind: string
  readonly target?: WorkbookRef
  readonly refs?: readonly WorkbookRef[]
  readonly message: string
}

export interface WorkbookChangeSummary {
  readonly kind: string
  readonly target?: WorkbookRef
  readonly message: string
}

export interface WorkbookUndoRef {
  readonly id: string
}

export interface WorkbookRunError {
  readonly code: string
  readonly message: string
}

export type WorkbookRunResult =
  | {
      readonly status: 'done'
      readonly changed: readonly WorkbookChangeSummary[]
      readonly checks: readonly WorkbookCheckResult[]
      readonly undo?: WorkbookUndoRef
    }
  | {
      readonly status: 'failed'
      readonly errors: readonly WorkbookRunError[]
      readonly checks: readonly WorkbookCheckResult[]
    }
