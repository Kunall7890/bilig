import { CellFlags } from '../../cell-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import type { FreshDirectAggregateFormulaBindingMember } from './formula-binding-service-types.js'
import type { InitialFormulaCellMembership } from './formula-initialization-membership.js'
import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'

type FormulaRowsByColumn = Map<number, readonly number[]>
type FormulaRowsBySheet = Map<number, FormulaRowsByColumn>

export interface InitialFreshAggregateFormulaDependencyIndex {
  readonly hasFormulaDependency: (member: FreshDirectAggregateFormulaBindingMember) => boolean
}

export function createInitialFreshAggregateFormulaDependencyIndex(args: {
  readonly pendingFormulaCells: InitialFormulaCellMembership | undefined
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'workbook'>
  readonly targetCellIndices: InitialFormulaCellIndexList
}): InitialFreshAggregateFormulaDependencyIndex {
  let formulaRowsBySheet: FormulaRowsBySheet | undefined

  const rowsBySheet = (): FormulaRowsBySheet => {
    if (formulaRowsBySheet !== undefined) {
      return formulaRowsBySheet
    }
    formulaRowsBySheet = buildFormulaRowsBySheet(args)
    return formulaRowsBySheet
  }

  return {
    hasFormulaDependency(member) {
      if (args.targetCellIndices.length === 0) {
        return false
      }
      const aggregateSheetId = args.state.workbook.getSheet(member.aggregateSheetName)?.id
      if (aggregateSheetId === undefined) {
        return false
      }
      const rowsByColumn = rowsBySheet().get(aggregateSheetId)
      if (rowsByColumn === undefined) {
        return false
      }
      return rowsByColumnIntersectsRange(rowsByColumn, member)
    },
  }
}

function buildFormulaRowsBySheet(args: {
  readonly pendingFormulaCells: InitialFormulaCellMembership | undefined
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'workbook'>
  readonly targetCellIndices: InitialFormulaCellIndexList
}): FormulaRowsBySheet {
  const rowsBySheet = new Map<number, Map<number, number[]>>()
  const cellStore = args.state.workbook.cellStore
  for (let index = 0; index < args.targetCellIndices.length; index += 1) {
    const formulaCellIndex = args.targetCellIndices[index]!
    const isFormulaCell =
      (args.pendingFormulaCells?.has(formulaCellIndex) ?? false) ||
      args.state.formulas.has(formulaCellIndex) ||
      ((cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.HasFormula) !== 0
    if (!isFormulaCell) {
      continue
    }
    const sheetId = cellStore.sheetIds[formulaCellIndex]
    const row = cellStore.rows[formulaCellIndex]
    const col = cellStore.cols[formulaCellIndex]
    if (sheetId === undefined || row === undefined || col === undefined) {
      continue
    }
    let rowsByColumn = rowsBySheet.get(sheetId)
    if (rowsByColumn === undefined) {
      rowsByColumn = new Map()
      rowsBySheet.set(sheetId, rowsByColumn)
    }
    let rows = rowsByColumn.get(col)
    if (rows === undefined) {
      rows = []
      rowsByColumn.set(col, rows)
    }
    rows.push(row)
  }
  for (const rowsByColumn of rowsBySheet.values()) {
    for (const rows of rowsByColumn.values()) {
      rows.sort((left, right) => left - right)
    }
  }
  return rowsBySheet
}

function rowsByColumnIntersectsRange(rowsByColumn: FormulaRowsByColumn, member: FreshDirectAggregateFormulaBindingMember): boolean {
  const colSpan = member.aggregateColEnd - member.aggregateColStart + 1
  if (colSpan <= rowsByColumn.size) {
    for (let col = member.aggregateColStart; col <= member.aggregateColEnd; col += 1) {
      const rows = rowsByColumn.get(col)
      if (rows !== undefined && sortedRowsIntersectRange(rows, member.aggregateRowStart, member.aggregateRowEnd)) {
        return true
      }
    }
    return false
  }

  for (const [col, rows] of rowsByColumn) {
    if (
      member.aggregateColStart <= col &&
      col <= member.aggregateColEnd &&
      sortedRowsIntersectRange(rows, member.aggregateRowStart, member.aggregateRowEnd)
    ) {
      return true
    }
  }
  return false
}

function sortedRowsIntersectRange(rows: readonly number[], start: number, end: number): boolean {
  let lower = 0
  let upper = rows.length
  while (lower < upper) {
    const mid = (lower + upper) >>> 1
    if (rows[mid]! < start) {
      lower = mid + 1
    } else {
      upper = mid
    }
  }
  return lower < rows.length && rows[lower]! <= end
}
