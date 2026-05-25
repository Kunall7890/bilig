import { createWorkbookRunAdapter, type SpreadsheetEngine } from '@bilig/core'
import {
  describeRunResult,
  runWorkbookPlan,
  type EngineOp,
  type WorkbookPlanData,
  type WorkbookRunResult,
  type WorkbookRunResultDescription,
} from '@bilig/workbook'
import type { WorkbookChangeUndoBundle } from '@bilig/zero-sync'

function engineOpsUndoBundle(ops: readonly EngineOp[] | undefined): WorkbookChangeUndoBundle | null {
  if (ops === undefined || ops.length === 0) {
    return null
  }
  return {
    kind: 'engineOps',
    ops: structuredClone([...ops]),
  }
}

export function workbookPlanRunAppliedOps(result: WorkbookRunResult): readonly EngineOp[] {
  return result.apply?.appliedOps ?? []
}

export function workbookPlanRunUndoBundle(result: WorkbookRunResult): WorkbookChangeUndoBundle | null {
  return engineOpsUndoBundle(result.undo?.ops)
}

export function workbookPlanRunResultProof(result: WorkbookRunResult): WorkbookRunResultDescription {
  return describeRunResult(result)
}

export async function runStrictWorkbookPlanData(
  engine: SpreadsheetEngine,
  plan: WorkbookPlanData,
  baseRevision = 0,
): Promise<WorkbookRunResult> {
  const result = await runWorkbookPlan(plan, createWorkbookRunAdapter(engine, { baseRevision }), {
    strict: true,
    expectedBaseRevision: baseRevision,
  })
  if (result.status === 'failed') {
    const undoOps = result.undo?.ops
    if (undoOps !== undefined && undoOps.length > 0) {
      engine.applyOps(undoOps, { source: 'restore', trusted: true })
    }
  }
  return result
}
