import type { EdgeArena } from '../../edge-arena.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { EngineRuntimeState, RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../runtime-state.js'

export function rebuildDirectAggregateFormulaDependenciesInPlace(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly edgeArena: EdgeArena
  readonly cellIndex: number
  readonly formula: RuntimeFormula
  readonly directAggregate: RuntimeDirectAggregateDescriptor
  readonly removeReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void
}): void {
  const formulaEntity = makeCellEntity(args.cellIndex)
  args.edgeArena.readView(args.formula.dependencyEntities).forEach((entity) => {
    args.removeReverseEdge(entity, formulaEntity)
  })

  const sheet = args.state.workbook.getSheet(args.directAggregate.sheetName)
  const dependencyIndices: number[] = []
  const dependencyEntities: number[] = []
  if (sheet) {
    for (let row = args.directAggregate.rowStart; row <= args.directAggregate.rowEnd; row += 1) {
      for (let col = args.directAggregate.col; col <= args.directAggregate.colEnd; col += 1) {
        const dependencyCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
        if (dependencyCellIndex === -1 || (args.state.workbook.cellStore.formulaIds[dependencyCellIndex] ?? 0) === 0) {
          continue
        }
        const dependencyEntity = makeCellEntity(dependencyCellIndex)
        dependencyIndices.push(dependencyCellIndex)
        dependencyEntities.push(dependencyEntity)
        args.appendKnownUniqueReverseEdge(dependencyEntity, formulaEntity)
      }
    }
  }

  args.formula.dependencyIndices = Uint32Array.from(dependencyIndices)
  args.formula.dependencyEntities = args.edgeArena.replace(args.formula.dependencyEntities, Uint32Array.from(dependencyEntities))
}
