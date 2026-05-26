import type { CodexDynamicToolCallResult, WorkbookAgentCommand, WorkbookAgentCommandBundle } from '@bilig/agent-api'
import {
  buildMutationReceipt,
  buildWorkbookAgentVerificationReport,
  selectWorkbookAgentMutationReceiptVerificationRanges,
  type WorkbookAgentToolStageContext,
  type WorkbookToolMutationReceipt,
} from './workbook-agent-mutation-receipt.js'
import type { WorkbookAgentMutationProofContext } from './workbook-agent-mutation-proof.js'
import { stringifyJson, textToolResult, type WorkbookAgentStageCommandResult } from './workbook-agent-tool-shared.js'

export interface WorkbookAgentVisibleCommitBarrierOutcome {
  readonly mutationExecuted: boolean
  readonly verificationComplete: boolean
  readonly status: WorkbookToolMutationReceipt['status']
  readonly summary: string
  readonly mutationReceipt: WorkbookToolMutationReceipt
  readonly appliedRevision: number | null
}

function summarizeWorkbookAgentVisibleCommitBarrierOutcome(input: {
  readonly bundle: WorkbookAgentCommandBundle
  readonly normalized: WorkbookAgentStageCommandResult
  readonly mutationReceipt: WorkbookToolMutationReceipt
}): string {
  const executionRecord = input.normalized.executionRecord
  if (executionRecord) {
    if (input.mutationReceipt.status === 'applied') {
      return `Applied workbook change set at revision r${String(executionRecord.appliedRevision)}: ${executionRecord.summary}`
    }
    const primaryWarning = input.mutationReceipt.warnings[0] ?? 'Mutation proof is incomplete.'
    return `Verification incomplete for workbook change set at revision r${String(executionRecord.appliedRevision)}: ${primaryWarning}`
  }
  if (input.normalized.disposition === 'queuedForTurnApply') {
    return `Queued workbook change set for turn apply and verification is incomplete: ${input.bundle.summary}`
  }
  return `Prepared workbook review item; the workbook is unchanged until this is applied: ${input.bundle.summary}`
}

export async function buildWorkbookAgentVisibleCommitBarrierOutcome(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly toolName: string
  readonly normalized: WorkbookAgentStageCommandResult
}): Promise<WorkbookAgentVisibleCommitBarrierOutcome> {
  const mutationReceipt = await buildMutationReceipt(input)
  return {
    mutationExecuted: input.normalized.executionRecord !== null,
    verificationComplete: mutationReceipt.status === 'applied',
    status: mutationReceipt.status,
    summary: summarizeWorkbookAgentVisibleCommitBarrierOutcome({
      bundle: input.normalized.bundle,
      normalized: input.normalized,
      mutationReceipt,
    }),
    mutationReceipt,
    appliedRevision: input.normalized.executionRecord?.appliedRevision ?? null,
  }
}

export async function stageWorkbookAgentVisibleCommitBarrierCommandResult(
  context: WorkbookAgentToolStageContext,
  command: WorkbookAgentCommand,
  toolName: string,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command)
  const normalized: WorkbookAgentStageCommandResult =
    'bundle' in result ? result : { bundle: result, executionRecord: null, disposition: 'reviewQueued' }
  const bundle = normalized.bundle
  const barrierOutcome = await buildWorkbookAgentVisibleCommitBarrierOutcome({
    context,
    toolName,
    normalized,
  })
  const mutationReceipt = barrierOutcome.mutationReceipt

  if (normalized.executionRecord) {
    const verification = await buildWorkbookAgentVerificationReport({
      context,
      revision: normalized.executionRecord.appliedRevision,
      ranges: selectWorkbookAgentMutationReceiptVerificationRanges(bundle),
    })
    return textToolResult(
      stringifyJson({
        applied: mutationReceipt.status === 'applied',
        mutationExecuted: barrierOutcome.mutationExecuted,
        verificationComplete: barrierOutcome.verificationComplete,
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: false,
        status: barrierOutcome.status,
        bundleId: bundle.id,
        summary: barrierOutcome.summary,
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
        mutationExecuted: barrierOutcome.mutationExecuted,
        verificationComplete: barrierOutcome.verificationComplete,
        staged: false,
        reviewQueued: false,
        queuedForTurnApply: true,
        status: barrierOutcome.status,
        bundleId: bundle.id,
        summary: barrierOutcome.summary,
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
      mutationExecuted: barrierOutcome.mutationExecuted,
      verificationComplete: barrierOutcome.verificationComplete,
      staged: true,
      reviewQueued: true,
      queuedForTurnApply: false,
      status: barrierOutcome.status,
      bundleId: bundle.id,
      summary: barrierOutcome.summary,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
      mutationReceipt,
    }),
  )
}
