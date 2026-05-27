import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import { ensureFormulaBindingDependencyBuildCapacity } from './formula-binding-dependency-build-capacity.js'
import { buildDirectScalarDescriptor, type DirectScalarWorkbook } from './formula-binding-direct-scalar.js'
import { appendFreshFormulaDependencyReverseEdges } from './formula-binding-install.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import { makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import { getFormulaBindingReverseEdgeSlice, setFormulaBindingReverseEdgeSlice } from './formula-binding-reverse-edges.js'
import type {
  CreateEngineFormulaBindingServiceArgs,
  FreshDirectScalarFormulaBindingInput,
  FreshDirectScalarFormulaBindingMember,
} from './formula-binding-service-types.js'
import type { RuntimeDirectScalarDescriptor, RuntimeDirectScalarOperand, RuntimeFormula } from '../runtime-state.js'

const EMPTY_U32 = new Uint32Array(0)

interface FreshDirectScalarDependencySlots {
  readonly dependencyIndex0: number
  readonly dependencyIndex1: number
  readonly dependencyIndexCount: number
  readonly dependencyEntity0: number
  readonly dependencyEntity1: number
  readonly dependencyEntityCount: number
}

interface FreshDirectScalarRunSlots {
  readonly dependencyIndex0: Uint32Array
  readonly dependencyIndex1: Uint32Array
  readonly dependencyIndexCounts: Uint8Array
  readonly dependencyEntity0: Uint32Array
  readonly dependencyEntity1: Uint32Array
  readonly dependencyEntityCounts: Uint8Array
}

export function bindFreshDirectScalarFormulaRun(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly edgeArena: EdgeArena
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void
  readonly run: FreshDirectScalarFormulaBindingInput
}): void {
  if ('member' in args.run) {
    assertFreshDirectScalarFormulaCell(args.serviceArgs, args.run.cellIndex)
    bindFreshDirectScalarFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.appendKnownUniqueReverseEdge,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      args.run.cellIndex,
      args.run.member,
    )
    return
  }

  if (args.run.cellIndices.length !== args.run.members.length) {
    throw new Error('Expected fresh direct scalar formula cell index count to match member count')
  }

  const run = args.run
  if (
    tryBindFreshDirectScalarFormulaRunBulk({
      serviceArgs: args.serviceArgs,
      edgeArena: args.edgeArena,
      formulaMemberCounts: args.formulaMemberCounts,
      appendKnownUniqueReverseEdge: args.appendKnownUniqueReverseEdge,
      trackFormulaSheetIndexes: args.trackFormulaSheetIndexes,
      run,
    })
  ) {
    return
  }
  if (args.serviceArgs.state.counters) {
    addEngineCounter(args.serviceArgs.state.counters, 'freshDirectScalarBulkFallbacks')
  }

  for (let index = 0; index < args.run.members.length; index += 1) {
    const cellIndex = args.run.cellIndices[index]!
    assertFreshDirectScalarFormulaCell(args.serviceArgs, cellIndex)
    bindFreshDirectScalarFormulaMember(
      args.serviceArgs,
      args.edgeArena,
      args.formulaMemberCounts,
      args.appendKnownUniqueReverseEdge,
      args.trackFormulaSheetIndexes,
      args.run.sheetId,
      args.run.ownerSheetName,
      cellIndex,
      args.run.members[index]!,
    )
  }
}

function tryBindFreshDirectScalarFormulaRunBulk(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly edgeArena: EdgeArena
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void
  readonly trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void
  readonly run: Exclude<FreshDirectScalarFormulaBindingInput, { readonly member: FreshDirectScalarFormulaBindingMember }>
}): boolean {
  if (args.run.members.length === 0) {
    return true
  }
  const memberCount = args.run.members.length
  const directScalars: RuntimeDirectScalarDescriptor[] = []
  directScalars.length = memberCount
  const formulaEntities = new Uint32Array(memberCount)
  const dependencySlots: FreshDirectScalarRunSlots = {
    dependencyIndex0: new Uint32Array(memberCount),
    dependencyIndex1: new Uint32Array(memberCount),
    dependencyIndexCounts: new Uint8Array(memberCount),
    dependencyEntity0: new Uint32Array(memberCount),
    dependencyEntity1: new Uint32Array(memberCount),
    dependencyEntityCounts: new Uint8Array(memberCount),
  }
  let totalDependencyIndexCount = 0
  let totalDependencyEntityCount = 0
  const directScalarWorkbook = createFreshRunDirectScalarWorkbook(args.serviceArgs.state.workbook, args.run.sheetId)

  for (let index = 0; index < memberCount; index += 1) {
    const cellIndex = args.run.cellIndices[index]!
    const member = args.run.members[index]!
    assertFreshDirectScalarFormulaCell(args.serviceArgs, cellIndex)
    const directScalar = buildDirectScalarDescriptor({
      compiled: member.compiled,
      ownerSheetName: args.run.ownerSheetName,
      ownerSheetId: args.run.sheetId,
      workbook: directScalarWorkbook,
      ensureCellTracked: args.serviceArgs.ensureCellTracked,
      ensureCellTrackedByCoords: args.serviceArgs.ensureCellTrackedByCoords,
    })
    if (directScalar === undefined) {
      return false
    }
    const dependencyCounts = writeFreshDirectScalarDependencySlotsForRun(member.compiled, directScalar, dependencySlots, index)
    if (dependencyCounts < 0) {
      return false
    }
    directScalars[index] = directScalar
    formulaEntities[index] = makeCellEntity(cellIndex)
    totalDependencyIndexCount += dependencyCounts & 0b11
    totalDependencyEntityCount += dependencyCounts >> 2
  }

  const packedDependencyIndices = new Uint32Array(totalDependencyIndexCount)
  const packedDependencyEntities = new Uint32Array(totalDependencyEntityCount)
  let dependencyIndexOffset = 0
  let dependencyEntityOffset = 0
  for (let index = 0; index < memberCount; index += 1) {
    writePackedFreshDirectScalarDependencySlots(
      dependencySlots,
      index,
      packedDependencyIndices,
      dependencyIndexOffset,
      packedDependencyEntities,
      dependencyEntityOffset,
    )
    dependencyIndexOffset += dependencySlots.dependencyIndexCounts[index]!
    dependencyEntityOffset += dependencySlots.dependencyEntityCounts[index]!
  }

  const dependencyEntityBlock =
    totalDependencyEntityCount === 0 ? args.edgeArena.empty() : args.edgeArena.replace(args.edgeArena.empty(), packedDependencyEntities)

  dependencyIndexOffset = 0
  dependencyEntityOffset = 0
  for (let index = 0; index < memberCount; index += 1) {
    const cellIndex = args.run.cellIndices[index]!
    const member = args.run.members[index]!
    const dependencyIndexCount = dependencySlots.dependencyIndexCounts[index]!
    const dependencyEntityCount = dependencySlots.dependencyEntityCounts[index]!
    const runtimeFormula = makeFreshDirectScalarRuntimeFormula(
      cellIndex,
      member,
      directScalars[index]!,
      packedDependencyIndices.subarray(dependencyIndexOffset, dependencyIndexOffset + dependencyIndexCount),
      subSlice(dependencyEntityBlock, dependencyEntityOffset, dependencyEntityCount),
    )
    const formulaSlotId = args.serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
    runtimeFormula.formulaSlotId = formulaSlotId
    args.formulaMemberCounts.increment(args.run.sheetId, member.col)
    markFormulaCellBound(args.serviceArgs.state.workbook.cellStore, cellIndex, member.compiled.mode)
    args.trackFormulaSheetIndexes(cellIndex, args.run.ownerSheetName, member.compiled)
    dependencyIndexOffset += dependencyIndexCount
    dependencyEntityOffset += dependencyEntityCount
  }

  if (
    !tryInstallFreshDirectScalarBulkReverseEdges(
      args.serviceArgs,
      args.edgeArena,
      formulaEntities,
      dependencySlots.dependencyEntityCounts,
      packedDependencyEntities,
    )
  ) {
    dependencyEntityOffset = 0
    for (let index = 0; index < memberCount; index += 1) {
      const dependencyEntityCount = dependencySlots.dependencyEntityCounts[index]!
      appendFreshFormulaDependencyReverseEdges(
        packedDependencyEntities.subarray(dependencyEntityOffset, dependencyEntityOffset + dependencyEntityCount),
        formulaEntities[index]!,
        args.appendKnownUniqueReverseEdge,
      )
      dependencyEntityOffset += dependencyEntityCount
    }
  }

  if (args.serviceArgs.state.counters) {
    addEngineCounter(args.serviceArgs.state.counters, 'freshDirectScalarBulkRunBindings')
    addEngineCounter(args.serviceArgs.state.counters, 'freshDirectScalarBulkMembers', memberCount)
    addEngineCounter(args.serviceArgs.state.counters, 'freshDirectScalarFormulaObjectsMaterialized', memberCount)
  }
  return true
}

function createFreshRunDirectScalarWorkbook(
  workbook: CreateEngineFormulaBindingServiceArgs['state']['workbook'],
  ownerSheetId: number,
): DirectScalarWorkbook {
  const ownerSheet = workbook.getSheetById(ownerSheetId)
  const cellStore = workbook.cellStore
  const fastFreshOwnerCellIndexAt =
    ownerSheet && ownerSheet.structureVersion === 1
      ? (row: number, col: number): number | undefined => {
          const cellIndex = ownerSheet.grid.getPhysical(row, col)
          return cellIndex !== -1 &&
            cellStore.sheetIds[cellIndex] === ownerSheetId &&
            cellStore.rows[cellIndex] === row &&
            cellStore.cols[cellIndex] === col
            ? cellIndex
            : undefined
        }
      : undefined

  return {
    getSheet: (sheetName) => workbook.getSheet(sheetName),
    getFreshCellIndexAt(sheetId, row, col) {
      if (sheetId === ownerSheetId && fastFreshOwnerCellIndexAt) {
        return fastFreshOwnerCellIndexAt(row, col)
      }
      return workbook.getFreshCellIndexAt(sheetId, row, col)
    },
    getCellIndexAt(sheetId, row, col) {
      return workbook.getCellIndexAt(sheetId, row, col)
    },
  }
}

function assertFreshDirectScalarFormulaCell(serviceArgs: CreateEngineFormulaBindingServiceArgs, cellIndex: number): void {
  if ((serviceArgs.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0) {
    throw new Error('Expected fresh direct scalar formula cell')
  }
}

function bindFreshDirectScalarFormulaMember(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  edgeArena: EdgeArena,
  formulaMemberCounts: FormulaBindingMemberCounts,
  appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void,
  trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void,
  sheetId: number,
  ownerSheetName: string,
  cellIndex: number,
  member: FreshDirectScalarFormulaBindingMember,
): void {
  const directScalar = buildDirectScalarDescriptor({
    compiled: member.compiled,
    ownerSheetName,
    ownerSheetId: sheetId,
    workbook: serviceArgs.state.workbook,
    ensureCellTracked: serviceArgs.ensureCellTracked,
    ensureCellTrackedByCoords: serviceArgs.ensureCellTrackedByCoords,
  })
  if (directScalar === undefined) {
    throw new Error('Expected fresh direct scalar formula descriptor')
  }
  const dependencies = materializeFreshDirectScalarDependencySlots(member.compiled, directScalar)
  if (dependencies === undefined) {
    throw new Error('Expected fresh direct scalar dependencies')
  }
  const dependencyEntities = edgeArena.replaceSmall(
    edgeArena.empty(),
    freshDirectScalarSmallDependencyCount(dependencies.dependencyEntityCount),
    dependencies.dependencyEntity0,
    dependencies.dependencyEntity1,
  )
  const runtimeFormula = makeFreshDirectScalarRuntimeFormula(
    cellIndex,
    member,
    directScalar,
    materializeFreshDirectScalarDependencyIndices(dependencies),
    dependencyEntities,
  )
  const formulaEntity = makeCellEntity(cellIndex)
  const formulaSlotId = serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
  runtimeFormula.formulaSlotId = formulaSlotId
  formulaMemberCounts.increment(sheetId, member.col)
  markFormulaCellBound(serviceArgs.state.workbook.cellStore, cellIndex, member.compiled.mode)
  appendFreshDirectScalarDependencyReverseEdgeSlots(dependencies, formulaEntity, appendKnownUniqueReverseEdge)
  trackFormulaSheetIndexes(cellIndex, ownerSheetName, member.compiled)
  if (serviceArgs.state.counters) {
    addEngineCounter(serviceArgs.state.counters, 'freshDirectScalarFormulaObjectsMaterialized')
  }
}

function makeFreshDirectScalarRuntimeFormula(
  cellIndex: number,
  member: FreshDirectScalarFormulaBindingMember,
  directScalar: RuntimeDirectScalarDescriptor,
  dependencyIndices: Uint32Array,
  dependencyEntities: EdgeSlice,
): RuntimeFormula {
  return {
    cellIndex,
    formulaSlotId: 0,
    planId: 0,
    templateId: member.templateId,
    source: member.source,
    compiled: member.compiled,
    plan: makeUnmanagedCompiledPlan(member.source, member.compiled, member.templateId),
    dependencyIndices,
    dependencyEntities,
    rangeDependencies: EMPTY_U32,
    graphRangeDependencies: EMPTY_U32,
    runtimeProgram: EMPTY_U32,
    constants: member.compiled.constants,
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: member.compiled.constants.length,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar,
    directCriteria: undefined,
  }
}

function materializeFreshDirectScalarDependencySlots(
  compiled: FreshDirectScalarFormulaBindingMember['compiled'],
  directScalar: RuntimeDirectScalarDescriptor,
): FreshDirectScalarDependencySlots | undefined {
  if (
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return undefined
  }
  let dependencyIndex0 = 0
  let dependencyIndex1 = 0
  let dependencyIndexCount = 0
  let dependencyEntity0 = 0
  let dependencyEntity1 = 0
  let dependencyEntityCount = 0
  const appendOperand = (operand: RuntimeDirectScalarOperand): boolean => {
    if (operand.kind === 'literal-number') {
      return true
    }
    if (operand.kind === 'error' || dependencyEntityCount >= 2) {
      return false
    }
    const cellIndex = operand.cellIndex
    if (dependencyIndexCount === 0) {
      dependencyIndex0 = cellIndex
      dependencyIndexCount = 1
    } else if (dependencyIndex0 !== cellIndex && dependencyIndexCount === 1) {
      dependencyIndex1 = cellIndex
      dependencyIndexCount = 2
    } else if (dependencyIndex0 !== cellIndex && dependencyIndex1 !== cellIndex) {
      return false
    }
    const entity = makeCellEntity(cellIndex)
    if (dependencyEntityCount === 0) {
      dependencyEntity0 = entity
    } else {
      dependencyEntity1 = entity
    }
    dependencyEntityCount += 1
    return true
  }
  const matched =
    directScalar.kind === 'abs'
      ? appendOperand(directScalar.operand)
      : appendOperand(directScalar.left) && appendOperand(directScalar.right)
  if (!matched) {
    return undefined
  }
  return {
    dependencyIndex0,
    dependencyIndex1,
    dependencyIndexCount,
    dependencyEntity0,
    dependencyEntity1,
    dependencyEntityCount,
  }
}

function materializeFreshDirectScalarDependencyIndices(slots: FreshDirectScalarDependencySlots): Uint32Array {
  if (slots.dependencyIndexCount === 0) {
    return EMPTY_U32
  }
  if (slots.dependencyIndexCount === 1) {
    const dependencyIndices = new Uint32Array(1)
    dependencyIndices[0] = slots.dependencyIndex0
    return dependencyIndices
  }
  const dependencyIndices = new Uint32Array(2)
  dependencyIndices[0] = slots.dependencyIndex0
  dependencyIndices[1] = slots.dependencyIndex1
  return dependencyIndices
}

function writeFreshDirectScalarDependencySlotsForRun(
  compiled: FreshDirectScalarFormulaBindingMember['compiled'],
  directScalar: RuntimeDirectScalarDescriptor,
  target: FreshDirectScalarRunSlots,
  index: number,
): number {
  if (
    compiled.symbolicRanges.length !== 0 ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0
  ) {
    return -1
  }

  let dependencyIndex0 = 0
  let dependencyIndex1 = 0
  let dependencyIndexCount = 0
  let dependencyEntity0 = 0
  let dependencyEntity1 = 0
  let dependencyEntityCount = 0
  const operandCount = directScalar.kind === 'abs' ? 1 : 2
  for (let operandIndex = 0; operandIndex < operandCount; operandIndex += 1) {
    const operand = directScalar.kind === 'abs' ? directScalar.operand : operandIndex === 0 ? directScalar.left : directScalar.right
    if (operand.kind === 'literal-number') {
      continue
    }
    if (operand.kind === 'error' || dependencyEntityCount >= 2) {
      return -1
    }
    const cellIndex = operand.cellIndex
    if (dependencyIndexCount === 0) {
      dependencyIndex0 = cellIndex
      dependencyIndexCount = 1
    } else if (dependencyIndex0 !== cellIndex && dependencyIndexCount === 1) {
      dependencyIndex1 = cellIndex
      dependencyIndexCount = 2
    } else if (dependencyIndex0 !== cellIndex && dependencyIndex1 !== cellIndex) {
      return -1
    }
    const entity = makeCellEntity(cellIndex)
    if (dependencyEntityCount === 0) {
      dependencyEntity0 = entity
    } else {
      dependencyEntity1 = entity
    }
    dependencyEntityCount += 1
  }

  target.dependencyIndex0[index] = dependencyIndex0
  target.dependencyIndex1[index] = dependencyIndex1
  target.dependencyIndexCounts[index] = dependencyIndexCount
  target.dependencyEntity0[index] = dependencyEntity0
  target.dependencyEntity1[index] = dependencyEntity1
  target.dependencyEntityCounts[index] = dependencyEntityCount
  return dependencyIndexCount | (dependencyEntityCount << 2)
}

function freshDirectScalarSmallDependencyCount(count: number): 0 | 1 | 2 {
  if (count === 0) {
    return 0
  }
  if (count === 1) {
    return 1
  }
  if (count === 2) {
    return 2
  }
  throw new Error('Expected fresh direct scalar dependency count to fit a small edge slice')
}

function appendFreshDirectScalarDependencyReverseEdgeSlots(
  slots: FreshDirectScalarDependencySlots,
  formulaEntity: number,
  appendKnownUniqueReverseEdge: (entityId: number, dependentEntityId: number) => void,
): void {
  if (slots.dependencyEntityCount === 0) {
    return
  }
  appendKnownUniqueReverseEdge(slots.dependencyEntity0, formulaEntity)
  if (slots.dependencyEntityCount > 1 && slots.dependencyEntity1 !== slots.dependencyEntity0) {
    appendKnownUniqueReverseEdge(slots.dependencyEntity1, formulaEntity)
  }
}

function writePackedFreshDirectScalarDependencySlots(
  slots: FreshDirectScalarRunSlots,
  index: number,
  packedDependencyIndices: Uint32Array,
  dependencyIndexOffset: number,
  packedDependencyEntities: Uint32Array,
  dependencyEntityOffset: number,
): void {
  const dependencyIndexCount = slots.dependencyIndexCounts[index]!
  if (dependencyIndexCount > 0) {
    packedDependencyIndices[dependencyIndexOffset] = slots.dependencyIndex0[index]!
  }
  if (dependencyIndexCount > 1) {
    packedDependencyIndices[dependencyIndexOffset + 1] = slots.dependencyIndex1[index]!
  }
  const dependencyEntityCount = slots.dependencyEntityCounts[index]!
  if (dependencyEntityCount > 0) {
    packedDependencyEntities[dependencyEntityOffset] = slots.dependencyEntity0[index]!
  }
  if (dependencyEntityCount > 1) {
    packedDependencyEntities[dependencyEntityOffset + 1] = slots.dependencyEntity1[index]!
  }
}

function subSlice(slice: EdgeSlice, offset: number, len: number): EdgeSlice {
  if (len === 0) {
    return EMPTY_EDGE_SLICE
  }
  return {
    ptr: slice.ptr + offset,
    len,
    cap: len,
  }
}

const EMPTY_EDGE_SLICE: EdgeSlice = { ptr: -1, len: 0, cap: 0 }

function tryInstallFreshDirectScalarBulkReverseEdges(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  edgeArena: EdgeArena,
  formulaEntities: Uint32Array,
  dependencyEntityCounts: Uint8Array,
  packedDependencyEntities: Uint32Array,
): boolean {
  if (packedDependencyEntities.length === 0) {
    return true
  }

  let epoch = serviceArgs.getDependencyBuildEpoch() + 1
  let seenDependencyCells = serviceArgs.getDependencyBuildSeen()
  if (epoch === 0xffff_ffff) {
    epoch = 1
    seenDependencyCells.fill(0)
  }
  serviceArgs.setDependencyBuildEpoch(epoch)

  for (let offset = 0; offset < packedDependencyEntities.length; offset += 1) {
    const dependencyEntity = packedDependencyEntities[offset]!
    if (dependencyEntity >= seenDependencyCells.length) {
      ensureFormulaBindingDependencyBuildCapacity(serviceArgs, dependencyEntity + 1, 0)
      seenDependencyCells = serviceArgs.getDependencyBuildSeen()
    }
    if (seenDependencyCells[dependencyEntity] === epoch) {
      return false
    }
    if (getFormulaBindingReverseEdgeSlice(serviceArgs.reverseState, dependencyEntity) !== undefined) {
      return false
    }
    seenDependencyCells[dependencyEntity] = epoch
  }

  const reverseBlock = edgeArena.replaceWith(edgeArena.empty(), packedDependencyEntities.length, (buffer, ptr) => {
    let dependencyEntityOffset = 0
    for (let index = 0; index < formulaEntities.length; index += 1) {
      const formulaEntity = formulaEntities[index]!
      const dependencyEntityCount = dependencyEntityCounts[index]!
      for (let offset = 0; offset < dependencyEntityCount; offset += 1) {
        buffer[ptr + dependencyEntityOffset + offset] = formulaEntity
      }
      dependencyEntityOffset += dependencyEntityCount
    }
  })
  for (let offset = 0; offset < packedDependencyEntities.length; offset += 1) {
    setFormulaBindingReverseEdgeSlice(serviceArgs.reverseState, packedDependencyEntities[offset]!, subSlice(reverseBlock, offset, 1))
  }
  if (serviceArgs.state.counters) {
    addEngineCounter(serviceArgs.state.counters, 'freshDirectScalarBulkReverseEdgeSlices', packedDependencyEntities.length)
  }
  return true
}
