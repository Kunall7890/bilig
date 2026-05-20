import type { FormulaNode } from '@bilig/formula'
import { FormulaMode } from '@bilig/protocol'
import type { EdgeSlice } from '../../edge-arena.js'
import type { FormulaInstanceSnapshot } from '../../formula/formula-instance-table.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { markFormulaCellBound } from './formula-binding-cell-flags.js'
import type { CachedFormulaInitializationRef, EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'

const EMPTY_U32 = new Uint32Array(0)
const EMPTY_EDGE_SLICE: EdgeSlice = { ptr: -1, len: 0, cap: 0 }
const CACHED_FORMULA_PLACEHOLDER_AST: FormulaNode = { kind: 'NumberLiteral', value: 0 }

function createCachedImportedRuntimeFormula(cellIndex: number, source: string): RuntimeFormula {
  const compiled: RuntimeFormula['compiled'] = {
    id: 0,
    source,
    mode: FormulaMode.JsOnly,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    maxStackDepth: 0,
    ast: CACHED_FORMULA_PLACEHOLDER_AST,
    optimizedAst: CACHED_FORMULA_PLACEHOLDER_AST,
    deps: [],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan: [],
    program: EMPTY_U32,
    constants: new Float64Array(0),
    symbolicRefs: [],
    symbolicRanges: [],
    symbolicStrings: [],
  }
  const plan = {
    id: 0,
    source,
    compiled,
  }
  return {
    cellIndex,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source,
    compiled,
    plan,
    dependencyIndices: EMPTY_U32,
    dependencyEntities: EMPTY_EDGE_SLICE,
    rangeDependencies: EMPTY_U32,
    graphRangeDependencies: EMPTY_U32,
    runtimeProgram: EMPTY_U32,
    constants: new Float64Array(0),
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
    preserveCachedValueOnFullRecalc: true,
    inlineScalarFastPlanKind: undefined,
    inlineScalarPlanCellIndices: undefined,
  }
}

export function initializeCachedFormulaSourcesAtNow(args: {
  readonly serviceArgs: EngineFormulaInitializationServiceArgs
  readonly refs: readonly CachedFormulaInitializationRef[]
  readonly potentialNewCells?: number
  readonly resolveSheetName: (sheetId: number) => string
}): void {
  const { serviceArgs, refs, potentialNewCells, resolveSheetName } = args
  if (refs.length === 0) {
    return
  }
  serviceArgs.beginMutationCollection()
  serviceArgs.checkEvaluationBudget()
  const reservedNewCells = potentialNewCells ?? refs.length
  const hadExistingFormulas = serviceArgs.state.formulas.size > 0
  serviceArgs.state.workbook.cellStore.ensureCapacity(serviceArgs.state.workbook.cellStore.size + reservedNewCells)
  serviceArgs.ensureRecalcScratchCapacity(serviceArgs.state.workbook.cellStore.size + reservedNewCells + 1)
  const formulaInstances: FormulaInstanceSnapshot[] | undefined =
    hadExistingFormulas || serviceArgs.hydrateFreshFormulaInstances === undefined ? undefined : []
  serviceArgs.state.workbook.withBatchedColumnVersionUpdates(() => {
    for (let index = 0; index < refs.length; index += 1) {
      serviceArgs.checkEvaluationBudget()
      const ref = refs[index]!
      const cellIndex = ref.cellIndex ?? serviceArgs.ensureCellTrackedByCoords(ref.sheetId, ref.row, ref.col)
      const runtimeFormula = createCachedImportedRuntimeFormula(cellIndex, ref.source)
      const formulaSlotId = serviceArgs.state.formulas.set(cellIndex, runtimeFormula)
      runtimeFormula.formulaSlotId = formulaSlotId
      markFormulaCellBound(serviceArgs.state.workbook.cellStore, cellIndex, FormulaMode.JsOnly)
      serviceArgs.writeHydratedFormulaValue(cellIndex, ref.value)
      if (formulaInstances) {
        formulaInstances.push({
          cellIndex,
          sheetName: resolveSheetName(ref.sheetId),
          row: ref.row,
          col: ref.col,
          source: ref.source,
        })
      }
    }
  })
  if (formulaInstances) {
    serviceArgs.hydrateFreshFormulaInstances?.(formulaInstances)
  } else {
    serviceArgs.deferFormulaInstanceTableRebuild?.()
  }
  const lastMetrics = serviceArgs.state.getLastMetrics()
  serviceArgs.state.setLastMetrics({
    ...lastMetrics,
    batchId: lastMetrics.batchId + 1,
    changedInputCount: 0,
    compileMs: lastMetrics.compileMs,
    recalcMs: 0,
  })
}
