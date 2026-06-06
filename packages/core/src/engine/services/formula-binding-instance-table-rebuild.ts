import type { FormulaInstanceSnapshot, FormulaInstanceTable } from '../../formula/formula-instance-table.js'
import type { FormulaOwnerPosition } from './formula-binding-service-types.js'

export interface FormulaBindingInstanceTableRebuildController {
  readonly deferRebuildNow: () => void
  readonly clearRebuildNow: () => void
  readonly recordIfMissingOrRebuildingNow: (
    cellIndex: number,
    source: string,
    templateId: number | undefined,
    ownerPosition?: FormulaOwnerPosition,
  ) => void
  readonly upsertFreshFormulaInstancesNow: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly hydrateFreshFormulaInstancesNow: (records: readonly FormulaInstanceSnapshot[]) => void
  readonly exportFormulaInstancesNow: () => FormulaInstanceSnapshot[]
}

export function createFormulaBindingInstanceTableRebuildController(args: {
  readonly formulaInstances: FormulaInstanceTable
  readonly rebuildFormulaInstancesNow: () => void
  readonly recordFormulaInstanceNow: (
    cellIndex: number,
    source: string,
    templateId: number | undefined,
    ownerPosition?: FormulaOwnerPosition,
  ) => void
}): FormulaBindingInstanceTableRebuildController {
  let needsRebuild = false

  const ensureFreshNow = (): void => {
    if (!needsRebuild) {
      return
    }
    needsRebuild = false
    args.rebuildFormulaInstancesNow()
  }

  return {
    deferRebuildNow() {
      args.formulaInstances.clear()
      needsRebuild = true
    },
    clearRebuildNow() {
      needsRebuild = false
    },
    recordIfMissingOrRebuildingNow(cellIndex, source, templateId, ownerPosition) {
      if (needsRebuild || !args.formulaInstances.get(cellIndex)) {
        args.recordFormulaInstanceNow(cellIndex, source, templateId, ownerPosition)
      }
    },
    upsertFreshFormulaInstancesNow(records) {
      if (!needsRebuild) {
        args.formulaInstances.upsertMany(records)
      }
    },
    hydrateFreshFormulaInstancesNow(records) {
      needsRebuild = false
      args.formulaInstances.hydrate(records)
    },
    exportFormulaInstancesNow() {
      ensureFreshNow()
      return args.formulaInstances.list()
    },
  }
}
