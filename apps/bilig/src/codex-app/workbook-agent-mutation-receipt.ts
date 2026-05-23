import type {
  CodexDynamicToolCallResult,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
  WorkbookAgentPreviewRange,
} from '@bilig/agent-api'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import { toWorkbookAgentRangeRef } from './workbook-agent-range-chunks.js'
import { emptyWorkbookRenderedReadbackProof, type WorkbookRenderedReadbackProof } from './workbook-agent-rendered-readback.js'
import {
  buildWorkbookAgentVerificationReport,
  buildWorkbookAuthoritativeReadbackProof,
  buildWorkbookRecalculationProof,
  buildWorkbookRenderedReadbackProof,
  resolveWorkbookMutationUndoStatus,
  type WorkbookAuthoritativeReadbackProof,
  type WorkbookMutationUndoProof,
  type WorkbookRecalculationProof,
  type WorkbookSemanticReadbackProof,
  type WorkbookAgentMutationProofContext,
} from './workbook-agent-mutation-proof.js'
import { stringifyJson, textToolResult, type WorkbookAgentStageCommandResult } from './workbook-agent-tool-shared.js'

const MAX_VERIFICATION_RANGES = 3

export { buildWorkbookAgentVerificationReport } from './workbook-agent-mutation-proof.js'

export interface WorkbookAgentMutationReceiptRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
  readonly role: WorkbookAgentPreviewRange['role']
  readonly kind: 'values' | 'formulas' | 'formats' | 'tables' | 'objects' | 'selection' | 'sheet'
}

export interface WorkbookToolMutationReceipt {
  readonly toolName: string
  readonly status: 'applied' | 'staged' | 'queued' | 'failed' | 'verification_incomplete'
  readonly revision: {
    readonly before: number | null
    readonly after: number | null
  }
  readonly affectedRanges: readonly WorkbookAgentMutationReceiptRange[]
  readonly recalculation: WorkbookRecalculationProof
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
  readonly renderedReadback: WorkbookRenderedReadbackProof
  readonly semanticReadback: WorkbookSemanticReadbackProof
  readonly undo: WorkbookMutationUndoProof
  readonly warnings: readonly string[]
}

export interface WorkbookAgentToolStageContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<WorkbookAgentCommandBundle | WorkbookAgentStageCommandResult>
}

function commandKind(command: WorkbookAgentCommand): WorkbookAgentMutationReceiptRange['kind'] {
  switch (command.kind) {
    case 'writeRange':
    case 'clearRange':
    case 'fillRange':
    case 'copyRange':
    case 'moveRange':
      return 'values'
    case 'setRangeFormulas':
      return 'formulas'
    case 'formatRange':
    case 'setDataValidation':
    case 'clearDataValidation':
    case 'upsertConditionalFormat':
    case 'deleteConditionalFormat':
    case 'setSheetProtection':
    case 'clearSheetProtection':
    case 'upsertRangeProtection':
    case 'deleteRangeProtection':
    case 'upsertCommentThread':
    case 'deleteCommentThread':
    case 'upsertNote':
    case 'deleteNote':
      return 'formats'
    case 'upsertTable':
    case 'deleteTable':
    case 'upsertPivotTable':
    case 'deletePivotTable':
      return 'tables'
    case 'upsertDefinedName':
    case 'deleteDefinedName':
    case 'upsertChart':
    case 'deleteChart':
    case 'upsertImage':
    case 'deleteImage':
    case 'upsertShape':
    case 'deleteShape':
      return 'objects'
    case 'createSheet':
    case 'renameSheet':
    case 'deleteSheet':
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns':
    case 'setFreezePane':
    case 'setFilter':
    case 'clearFilter':
    case 'setSort':
    case 'clearSort':
    case 'updateRowMetadata':
    case 'updateColumnMetadata':
      return 'sheet'
    default: {
      const exhaustive: never = command
      return exhaustive
    }
  }
}

function receiptRanges(bundle: WorkbookAgentCommandBundle): readonly WorkbookAgentMutationReceiptRange[] {
  return bundle.affectedRanges.map((range) => ({
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
    role: range.role,
    kind: commandKind(bundle.commands[0] ?? ({ kind: 'clearRange', range } as WorkbookAgentCommand)),
  }))
}

function verificationRanges(bundle: WorkbookAgentCommandBundle): readonly CellRangeRef[] {
  return bundle.affectedRanges
    .filter((range) => range.role === 'target')
    .slice(0, MAX_VERIFICATION_RANGES)
    .map((range) => toWorkbookAgentRangeRef(range))
}

function buildAppliedMutationSummary(input: {
  readonly mutationReceipt: WorkbookToolMutationReceipt
  readonly appliedRevision: number
  readonly executionSummary: string
}): string {
  if (input.mutationReceipt.status === 'applied') {
    return `Applied workbook change set at revision r${String(input.appliedRevision)}: ${input.executionSummary}`
  }
  const primaryWarning = input.mutationReceipt.warnings[0] ?? 'Mutation proof is incomplete.'
  return `Verification incomplete for workbook change set at revision r${String(input.appliedRevision)}: ${primaryWarning}`
}

function buildWorkbookSemanticReadbackProof(input: {
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
  readonly renderedReadback: WorkbookRenderedReadbackProof
}): WorkbookSemanticReadbackProof {
  const requested = input.authoritativeReadback.requested || input.renderedReadback.requested
  if (!requested) {
    return {
      requested: false,
      matched: null,
      incompleteReason: input.authoritativeReadback.incompleteReason ?? input.renderedReadback.incompleteReason,
    }
  }
  const matched =
    input.authoritativeReadback.matched === true && (!input.renderedReadback.requested || input.renderedReadback.matched === true)
  return {
    requested,
    matched,
    incompleteReason:
      input.authoritativeReadback.matched !== true
        ? (input.authoritativeReadback.incompleteReason ?? 'Authoritative semantic readback did not match.')
        : input.renderedReadback.requested && input.renderedReadback.matched !== true
          ? (input.renderedReadback.incompleteReason ?? 'Rendered semantic readback did not match.')
          : null,
  }
}

export async function buildMutationReceipt(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly toolName: string
  readonly normalized: WorkbookAgentStageCommandResult
}): Promise<WorkbookToolMutationReceipt> {
  const { bundle, executionRecord } = input.normalized
  const ranges = executionRecord ? verificationRanges(bundle) : []
  const authoritativeReadback = executionRecord
    ? await buildWorkbookAuthoritativeReadbackProof({
        context: input.context,
        bundle,
        executionRecord,
        ranges,
      })
    : {
        requested: false,
        matched: null,
        ranges: [],
        mismatches: [],
        incompleteReason: 'Workbook mutation is not applied, so authoritative readback is not yet meaningful.',
      }
  const renderedReadback = executionRecord
    ? await buildWorkbookRenderedReadbackProof({
        context: input.context,
        appliedRevision: executionRecord.appliedRevision,
        ranges,
        authoritativeReadback,
      })
    : emptyWorkbookRenderedReadbackProof({
        requested: false,
        reason: 'Workbook mutation is not applied, so rendered readback is not yet meaningful.',
      })
  const semanticReadback = buildWorkbookSemanticReadbackProof({
    authoritativeReadback,
    renderedReadback,
  })
  const undo = await resolveWorkbookMutationUndoStatus({
    context: input.context,
    appliedRevision: executionRecord?.appliedRevision ?? null,
  })
  const recalculation = await buildWorkbookRecalculationProof({
    context: input.context,
    appliedRevision: executionRecord?.appliedRevision ?? null,
  })
  const warnings: string[] = []
  if (!executionRecord && input.normalized.disposition === 'queuedForTurnApply') {
    warnings.push(
      'Queued workbook change sets are not completed mutations. The assistant must wait for apply and verify before claiming success.',
    )
  }
  if (!executionRecord && input.normalized.disposition === 'reviewQueued') {
    warnings.push('Workbook change set is waiting for owner review and has not modified the workbook yet.')
  }
  if (executionRecord && authoritativeReadback.matched !== true) {
    warnings.push(authoritativeReadback.incompleteReason ?? 'Authoritative readback did not prove the mutation.')
  }
  if (executionRecord && renderedReadback.matched !== true) {
    warnings.push(renderedReadback.incompleteReason ?? 'Rendered readback did not prove the mutation.')
  }
  if (executionRecord && recalculation.matched !== true) {
    warnings.push(recalculation.incompleteReason ?? 'Workbook recalculation did not prove the mutation.')
  }
  if (executionRecord && !undo.available) {
    warnings.push(undo.reasonUnavailable ?? 'Undo status is unavailable.')
  }
  const hasAppliedProof =
    executionRecord !== null &&
    recalculation.requested &&
    recalculation.matched === true &&
    authoritativeReadback.requested &&
    authoritativeReadback.matched === true &&
    renderedReadback.requested &&
    renderedReadback.matched === true &&
    semanticReadback.requested &&
    semanticReadback.matched === true &&
    undo.available
  return {
    toolName: input.toolName,
    status: executionRecord
      ? hasAppliedProof
        ? 'applied'
        : 'verification_incomplete'
      : input.normalized.disposition === 'queuedForTurnApply'
        ? 'queued'
        : 'staged',
    revision: {
      before: bundle.baseRevision,
      after: executionRecord?.appliedRevision ?? null,
    },
    affectedRanges: receiptRanges(bundle),
    recalculation,
    authoritativeReadback,
    renderedReadback,
    semanticReadback,
    undo,
    warnings,
  }
}

export async function stageWorkbookAgentCommandResult(
  context: WorkbookAgentToolStageContext,
  command: WorkbookAgentCommand,
  toolName: string,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command)
  const normalized: WorkbookAgentStageCommandResult =
    'bundle' in result ? result : { bundle: result, executionRecord: null, disposition: 'reviewQueued' }
  const bundle = normalized.bundle
  const mutationReceipt = await buildMutationReceipt({
    context,
    toolName,
    normalized,
  })
  if (normalized.executionRecord) {
    const verification = await buildWorkbookAgentVerificationReport({
      context,
      revision: normalized.executionRecord.appliedRevision,
      ranges: verificationRanges(bundle),
    })
    return textToolResult(
      stringifyJson({
        applied: mutationReceipt.status === 'applied',
        mutationExecuted: true,
        verificationComplete: mutationReceipt.status === 'applied',
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: false,
        status: mutationReceipt.status,
        bundleId: bundle.id,
        summary: buildAppliedMutationSummary({
          mutationReceipt,
          appliedRevision: normalized.executionRecord.appliedRevision,
          executionSummary: normalized.executionRecord.summary,
        }),
        revision: normalized.executionRecord.appliedRevision,
        scope: normalized.executionRecord.scope,
        riskClass: normalized.executionRecord.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
        mutationReceipt,
        verification,
      }),
    )
  }
  if (normalized.disposition === 'queuedForTurnApply') {
    return textToolResult(
      stringifyJson({
        applied: false,
        mutationExecuted: false,
        verificationComplete: false,
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: true,
        status: 'queued',
        bundleId: bundle.id,
        summary: `Queued workbook change set for turn apply and verification is incomplete: ${bundle.summary}`,
        scope: bundle.scope,
        riskClass: bundle.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
        mutationReceipt,
      }),
    )
  }
  return textToolResult(
    stringifyJson({
      applied: false,
      mutationExecuted: false,
      verificationComplete: false,
      staged: true,
      reviewQueued: true,
      queuedForTurnApply: false,
      status: 'staged',
      bundleId: bundle.id,
      summary: `Prepared workbook review item; the workbook is unchanged until this is applied: ${bundle.summary}`,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
      mutationReceipt,
    }),
  )
}
