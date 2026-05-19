import { FormulaMode } from '@bilig/protocol'
import type { EdgeArena } from '../../edge-arena.js'
import { appendDirectAggregateColumnReverseEdges } from './formula-binding-dependency-helpers.js'
import type { FormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import { makeUnmanagedCompiledPlan } from './formula-binding-plan-helpers.js'
import type {
  CreateEngineFormulaBindingServiceArgs,
  FreshDirectAggregateFormulaBindingMember,
  FreshDirectAggregateFormulaBindingRun,
} from './formula-binding-service-types.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../runtime-state.js'

const EMPTY_U32 = new Uint32Array(0)

export function bindFreshDirectAggregateFormulaRun(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly edgeArena: EdgeArena
  readonly formulaMemberCounts: FormulaBindingMemberCounts
  readonly trackFormulaSheetIndexes: (
    cellIndex: number,
    ownerSheetName: string,
    compiled: Pick<RuntimeFormula['compiled'], 'deps' | 'parsedDeps'>,
  ) => void
  readonly run: FreshDirectAggregateFormulaBindingRun
}): void {
  if (args.run.cellIndices.length !== args.run.members.length) {
    throw new Error('Expected fresh direct aggregate formula cell index count to match member count')
  }
  for (let index = 0; index < args.run.members.length; index += 1) {
    const member = args.run.members[index]!
    const cellIndex = args.run.cellIndices[index]!
    if ((args.serviceArgs.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0) {
      throw new Error('Expected fresh direct aggregate formula cell')
    }
    if (
      member.aggregateRowStart > member.aggregateRowEnd ||
      member.aggregateColStart > member.aggregateColEnd ||
      (member.aggregateRowStart <= member.row &&
        member.row <= member.aggregateRowEnd &&
        member.aggregateColStart <= member.col &&
        member.col <= member.aggregateColEnd)
    ) {
      throw new Error('Expected non-recursive fresh direct aggregate formula')
    }
  }

  for (let index = 0; index < args.run.members.length; index += 1) {
    const member = args.run.members[index]!
    const cellIndex = args.run.cellIndices[index]!
    const directAggregate = buildFreshDirectAggregateDescriptor(args.serviceArgs, args.run.ownerSheetName, member)
    const runtimeFormula: RuntimeFormula = {
      cellIndex,
      formulaSlotId: 0,
      planId: 0,
      templateId: member.templateId,
      source: member.source,
      compiled: member.compiled,
      plan: makeUnmanagedCompiledPlan(member.source, member.compiled, member.templateId),
      dependencyIndices: EMPTY_U32,
      dependencyEntities: args.edgeArena.empty(),
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
      directAggregate,
      directScalar: undefined,
      directCriteria: undefined,
    }
    const formulaSlotId = args.serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
    runtimeFormula.formulaSlotId = formulaSlotId
    args.formulaMemberCounts.increment(args.run.sheetId, member.col)
    markFormulaCellBound(args.serviceArgs.state.workbook.cellStore, cellIndex, member.compiled.mode)
    appendDirectAggregateColumnReverseEdges(
      args.serviceArgs.reverseState.reverseAggregateColumnEdges,
      args.serviceArgs.state.workbook,
      directAggregate,
      cellIndex,
    )
    args.trackFormulaSheetIndexes(cellIndex, args.run.ownerSheetName, member.compiled)
    if (member.compiled.mode === FormulaMode.WasmFastPath && member.compiled.program.length > 0) {
      args.serviceArgs.scheduleWasmProgramSync()
    }
  }
}

function buildFreshDirectAggregateDescriptor(
  serviceArgs: CreateEngineFormulaBindingServiceArgs,
  ownerSheetName: string,
  member: FreshDirectAggregateFormulaBindingMember,
): RuntimeDirectAggregateDescriptor {
  return {
    regionId: serviceArgs.regionGraph.internSingleColumnRegion({
      sheetName: ownerSheetName,
      rowStart: member.aggregateRowStart,
      rowEnd: member.aggregateRowEnd,
      col: member.aggregateColStart,
    }),
    aggregateKind: member.aggregateKind,
    sheetName: ownerSheetName,
    rowStart: member.aggregateRowStart,
    rowEnd: member.aggregateRowEnd,
    col: member.aggregateColStart,
    colEnd: member.aggregateColEnd,
    length: (member.aggregateRowEnd - member.aggregateRowStart + 1) * (member.aggregateColEnd - member.aggregateColStart + 1),
    ...(member.resultOffset !== undefined ? { resultOffset: member.resultOffset } : {}),
  }
}
