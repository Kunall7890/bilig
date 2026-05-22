import type { CellRangeRef } from '@bilig/protocol'
import type { ZeroSyncService } from '../zero/service.js'
import { verifyWorkbookInvariants } from './workbook-agent-audit.js'
import { findWorkbookFormulaIssues } from './workbook-agent-comprehension.js'
import { inspectWorkbookRange, normalizeWorkbookAgentUiContext } from './workbook-agent-inspection.js'
import { countWorkbookAgentRangeCells, createWorkbookAgentRangeChunkPlan, toWorkbookAgentRangeRef } from './workbook-agent-range-chunks.js'
import {
  emptyWorkbookRenderedReadbackProof,
  selectWorkbookRenderedReadback,
  type WorkbookRenderedReadbackProof,
} from './workbook-agent-rendered-readback.js'
import { firstAuthoritativeRows, MAX_RECEIPT_READBACK_CELLS } from './workbook-agent-mutation-proof.js'
import type {
  WorkbookAgentMutationProofContext,
  WorkbookAuthoritativeReadbackProof,
  WorkbookMutationRecalculationProof,
  WorkbookMutationUndoProof,
} from './workbook-agent-mutation-proof-types.js'

function describeUnknownError(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message
  }
  if (typeof error === 'string' && error.trim().length > 0) {
    return error
  }
  return 'unknown error'
}

function asReadonlyRows(rows: readonly unknown[]): readonly (readonly unknown[])[] {
  return rows.filter((row): row is unknown[] => Array.isArray(row))
}

export async function resolveWorkbookMutationUndoStatus(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly appliedRevision: number | null
}): Promise<WorkbookMutationUndoProof> {
  if (input.appliedRevision === null) {
    return {
      available: false,
      token: null,
      reasonUnavailable: 'Workbook mutation has not been applied yet.',
      lookupFailed: false,
    }
  }
  let changes: Awaited<ReturnType<ZeroSyncService['listWorkbookChanges']>>
  try {
    changes = await input.context.zeroSyncService.listWorkbookChanges(input.context.documentId, 25)
  } catch (error) {
    return {
      available: false,
      token: null,
      reasonUnavailable: `Undo metadata lookup failed for applied revision r${String(input.appliedRevision)}: ${describeUnknownError(error)}`,
      lookupFailed: true,
    }
  }
  const matchingChange = changes.find((change) => change.revision === input.appliedRevision) ?? null
  if (matchingChange?.undoBundle) {
    return {
      available: true,
      token: `revision:${String(input.appliedRevision)}`,
      reasonUnavailable: null,
      lookupFailed: false,
    }
  }
  return {
    available: false,
    token: null,
    reasonUnavailable: 'No persisted undo metadata was returned for the applied revision.',
    lookupFailed: false,
  }
}

export async function resolveWorkbookMutationRecalculationStatus(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly appliedRevision: number | null
}): Promise<WorkbookMutationRecalculationProof> {
  if (input.appliedRevision === null) {
    return {
      requested: false,
      upToDate: null,
      appliedRevision: null,
      headRevision: null,
      calculatedRevision: null,
      lastMetrics: null,
      incompleteReason: 'Workbook mutation has not been applied yet.',
    }
  }
  try {
    return await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, (runtime) => {
      const requiredRevision = Math.max(runtime.headRevision, input.appliedRevision ?? 0)
      const upToDate = runtime.calculatedRevision >= requiredRevision
      return {
        requested: true,
        upToDate,
        appliedRevision: input.appliedRevision,
        headRevision: runtime.headRevision,
        calculatedRevision: runtime.calculatedRevision,
        lastMetrics: runtime.engine.getLastMetrics(),
        incompleteReason: upToDate
          ? null
          : `Workbook recalculation is stale: calculated revision r${String(runtime.calculatedRevision)} is behind required revision r${String(
              requiredRevision,
            )}.`,
      }
    })
  } catch (error) {
    return {
      requested: true,
      upToDate: false,
      appliedRevision: input.appliedRevision,
      headRevision: null,
      calculatedRevision: null,
      lastMetrics: null,
      incompleteReason: `Workbook recalculation proof failed for applied revision r${String(input.appliedRevision)}: ${describeUnknownError(
        error,
      )}`,
    }
  }
}

export async function buildWorkbookRenderedReadbackProof(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly appliedRevision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly authoritativeReadback: WorkbookAuthoritativeReadbackProof
}): Promise<WorkbookRenderedReadbackProof> {
  const range = input.ranges[0] ?? null
  if (!range) {
    return emptyWorkbookRenderedReadbackProof({
      requested: false,
      reason: 'No target cell range was available for rendered readback.',
    })
  }
  const nextChunk =
    countWorkbookAgentRangeCells(range) > MAX_RECEIPT_READBACK_CELLS
      ? (createWorkbookAgentRangeChunkPlan(range, MAX_RECEIPT_READBACK_CELLS).chunks[1] ?? null)
      : null
  const uiContext = await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, (runtime) =>
    normalizeWorkbookAgentUiContext(runtime, input.context.uiContext),
  )
  const authoritativeRows = firstAuthoritativeRows(input.authoritativeReadback)
  return selectWorkbookRenderedReadback({
    renderedContext: uiContext?.rendered,
    requestedRange: range,
    minRevision: input.appliedRevision,
    nextChunk,
    ...(authoritativeRows !== undefined ? { authoritativeRows } : {}),
  })
}

export async function buildWorkbookAgentVerificationReport(input: {
  readonly context: WorkbookAgentMutationProofContext
  readonly revision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly includeFormulaIssues?: boolean
  readonly includeInvariants?: boolean
}) {
  return await input.context.zeroSyncService.inspectWorkbook(input.context.documentId, async (runtime) => {
    const uiContext = normalizeWorkbookAgentUiContext(runtime, input.context.uiContext)
    const normalizedRanges = input.ranges.map((range) => toWorkbookAgentRangeRef(range))
    const authoritativeReadback = normalizedRanges.map((range) => inspectWorkbookRange(runtime, range))
    const renderedReadback = normalizedRanges.map((range) => {
      const authoritativeRange = inspectWorkbookRange(runtime, range)
      return selectWorkbookRenderedReadback({
        renderedContext: uiContext?.rendered,
        requestedRange: range,
        authoritativeRows: asReadonlyRows(authoritativeRange.rows),
        minRevision: input.revision,
      })
    })
    const formulaIssues =
      input.includeFormulaIssues === false
        ? null
        : findWorkbookFormulaIssues(runtime, {
            limit: 100,
          })
    const invariants = input.includeInvariants === false ? null : await verifyWorkbookInvariants(runtime, { roundTrip: true })
    return {
      appliedRevision: input.revision,
      recalculationStatus: {
        headRevision: runtime.headRevision,
        calculatedRevision: runtime.calculatedRevision,
        upToDate: runtime.calculatedRevision >= runtime.headRevision,
        lastMetrics: runtime.engine.getLastMetrics(),
      },
      authoritativeReadback,
      renderedReadback,
      formulaIssues,
      invariants,
    }
  })
}
