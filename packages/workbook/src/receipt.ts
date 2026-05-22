import type { WorkbookRef } from './find.js'
import type { WorkbookActionInput } from './input.js'
import type { WorkbookRevision } from './command.js'
import type { WorkbookChangeSummary, WorkbookUndoRef } from './result.js'

export type WorkbookReceiptProofStatus = 'passed' | 'failed' | 'skipped'

export type WorkbookReceiptProofKind =
  | 'preview'
  | 'apply'
  | 'authoritativeReadback'
  | 'renderedReadback'
  | 'semanticReadback'
  | 'recalculation'
  | 'undo'
  | 'check'
  | 'custom'

export const workbookReceiptProofKinds = Object.freeze([
  'preview',
  'apply',
  'authoritativeReadback',
  'renderedReadback',
  'semanticReadback',
  'recalculation',
  'undo',
  'check',
  'custom',
] satisfies readonly WorkbookReceiptProofKind[])

export function isWorkbookReceiptProofKind(value: unknown): value is WorkbookReceiptProofKind {
  return typeof value === 'string' && workbookReceiptProofKinds.some((kind) => kind === value)
}

export interface WorkbookReceiptProof {
  readonly kind: WorkbookReceiptProofKind
  readonly status: WorkbookReceiptProofStatus
  readonly message: string
  readonly revision?: WorkbookRevision
  readonly target?: WorkbookRef
  readonly data?: WorkbookActionInput
}

export interface WorkbookRenderedReceipt {
  readonly revision?: WorkbookRevision
  readonly diffs?: readonly WorkbookChangeSummary[]
  readonly message?: string
}

export interface WorkbookRuntimeReceipt {
  readonly appliedRevision?: WorkbookRevision
  readonly calculatedRevision?: WorkbookRevision
  readonly renderedRevision?: WorkbookRevision
  readonly rendered?: WorkbookRenderedReceipt
  readonly proof?: readonly WorkbookReceiptProof[]
  readonly warnings?: readonly string[]
}

export interface WorkbookRunReceipt {
  readonly commandId?: string
  readonly idempotencyKey?: string
  readonly modelName: string
  readonly actionName: string
  readonly baseRevision?: WorkbookRevision
  readonly appliedRevision?: WorkbookRevision
  readonly calculatedRevision?: WorkbookRevision
  readonly renderedRevision?: WorkbookRevision
  readonly rendered?: WorkbookRenderedReceipt
  readonly previewed: boolean
  readonly applied: boolean
  readonly verified: boolean
  readonly checkCount: number
  readonly passedCheckCount: number
  readonly failedCheckCount: number
  readonly unverifiedCheckCount: number
  readonly proof: readonly WorkbookReceiptProof[]
  readonly warnings?: readonly string[]
  readonly undo?: WorkbookUndoRef
}
