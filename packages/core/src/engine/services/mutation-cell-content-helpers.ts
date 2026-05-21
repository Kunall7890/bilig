import { parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import type { CellSnapshot } from '@bilig/protocol'

export function getMutationMatrixCell(matrix: readonly (readonly CellSnapshot[])[], rowIndex: number, colIndex: number): CellSnapshot {
  const row = matrix[rowIndex]
  if (row === undefined) {
    throw new RangeError(`Missing source row at index ${rowIndex}`)
  }
  const cell = row[colIndex]
  if (cell === undefined) {
    throw new RangeError(`Missing source cell at row ${rowIndex}, column ${colIndex}`)
  }
  return cell
}

export function isMutationStructuralInsertOp(op: EngineOp): op is Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }> {
  return op.kind === 'insertRows' || op.kind === 'insertColumns'
}

export function inverseMutationStructuralInsertOp(
  op: Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }>,
): Extract<EngineOp, { kind: 'deleteRows' | 'deleteColumns' }> {
  return op.kind === 'insertRows'
    ? { kind: 'deleteRows', sheetName: op.sheetName, start: op.start, count: op.count }
    : { kind: 'deleteColumns', sheetName: op.sheetName, start: op.start, count: op.count }
}

export function collectLiveCreatedSheetNames(
  existingSheetNames: Iterable<string>,
  ops: ReadonlyArray<{
    readonly kind: string
    readonly name?: string
    readonly oldName?: string
    readonly newName?: string
  }>,
): ReadonlySet<string> {
  const knownSheetNames = new Set(existingSheetNames)
  const liveCreatedSheetNames = new Set<string>()

  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index]
    if (!op) {
      continue
    }
    if (op.kind === 'upsertSheet') {
      const sheetName = op.name
      if (typeof sheetName === 'string' && !knownSheetNames.has(sheetName)) {
        liveCreatedSheetNames.add(sheetName)
      }
      if (typeof sheetName === 'string') {
        knownSheetNames.add(sheetName)
      }
      continue
    }
    if (op.kind !== 'renameSheet') {
      continue
    }
    const oldName = op.oldName
    const newName = op.newName
    if (typeof oldName === 'string' && typeof newName === 'string' && liveCreatedSheetNames.delete(oldName)) {
      liveCreatedSheetNames.add(newName)
    }
    if (typeof oldName === 'string') {
      knownSheetNames.delete(oldName)
    }
    if (typeof newName === 'string') {
      knownSheetNames.add(newName)
    }
  }

  return liveCreatedSheetNames
}

export function translateMutationFormulaForTarget(
  formula: string,
  sourceSheetName: string,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): string {
  if (sourceSheetName !== targetSheetName) {
    return formula
  }
  const source = parseCellAddress(sourceAddress, sourceSheetName)
  const target = parseCellAddress(targetAddress, targetSheetName)
  return translateFormulaReferences(formula, target.row - source.row, target.col - source.col)
}
