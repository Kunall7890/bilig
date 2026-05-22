import type { CellRangeRef, WorkbookPivotSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import type { WorkbookPivotRecord } from '../../workbook-store.js'

export type SourcefulPivotSnapshot = WorkbookPivotSnapshot & { source: CellRangeRef }
export type PivotUpsertOp = Extract<EngineOp, { kind: 'upsertPivotTable' }>

export function sourcefulPivotToUpsertOp(pivot: SourcefulPivotSnapshot): PivotUpsertOp {
  const cloned = structuredClone(pivot)
  return {
    kind: 'upsertPivotTable',
    ...cloned,
  }
}

export function pivotUpsertOpToRecord(op: PivotUpsertOp): WorkbookPivotRecord {
  const { kind, ...pivot } = op
  void kind
  return structuredClone(pivot)
}
