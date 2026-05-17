import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import type { CompiledFormula } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import { EdgeArena } from '../edge-arena.js'
import { SpreadsheetEngine } from '../engine.js'
import { createInitialRecalcMetrics } from '../engine/runtime-state.js'
import { createEngineMaintenanceService, type EngineMaintenanceService } from '../engine/services/maintenance-service.js'
import { FormulaTable } from '../formula-table.js'
import { RangeRegistry } from '../range-registry.js'
import { WorkbookStore } from '../workbook-store.js'

type MaintenanceServiceArgs = Parameters<typeof createEngineMaintenanceService>[0]
type MaintenanceServiceArgsOverrides = Omit<Partial<MaintenanceServiceArgs>, 'state' | 'reverseState'> & {
  readonly state?: Partial<MaintenanceServiceArgs['state']>
  readonly reverseState?: Partial<MaintenanceServiceArgs['reverseState']>
}

class ThrowingMaintenanceWorkbook extends WorkbookStore {
  override listDefinedNames() {
    throw new Error('rename boom')
  }

  override reset(_workbookName = 'Workbook'): void {
    throw new Error('reset boom')
  }
}

function isEngineMaintenanceService(value: unknown): value is EngineMaintenanceService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'estimatePotentialNewCells') === 'function' &&
    typeof Reflect.get(value, 'rewriteDefinedNamesForSheetRename') === 'function' &&
    typeof Reflect.get(value, 'resetWorkbook') === 'function'
  )
}

function getMaintenanceService(engine: SpreadsheetEngine): EngineMaintenanceService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const maintenance = Reflect.get(runtime, 'maintenance')
  if (!isEngineMaintenanceService(maintenance)) {
    throw new TypeError('Expected engine maintenance service')
  }
  return maintenance
}

function createMaintenanceServiceArgs(overrides: MaintenanceServiceArgsOverrides = {}): MaintenanceServiceArgs {
  const workbook = overrides.state?.workbook ?? new WorkbookStore('maintenance-service-stub')
  const state: MaintenanceServiceArgs['state'] = {
    workbook,
    formulas: new FormulaTable<CompiledFormula>(workbook.cellStore),
    ranges: new RangeRegistry(),
    entityVersions: new Map(),
    sheetDeleteVersions: new Map(),
    undoStack: [],
    redoStack: [],
    setSelection: () => undefined,
    setSyncState: () => undefined,
    getLastMetrics: () => ({ ...createInitialRecalcMetrics(), batchId: 'batch-1' }),
    setLastMetrics: () => undefined,
    ...overrides.state,
  }
  const reverseState: MaintenanceServiceArgs['reverseState'] = {
    reverseCellEdges: [],
    reverseRangeEdges: [],
    reverseDefinedNameEdges: new Map(),
    reverseTableEdges: new Map(),
    reverseSpillEdges: new Map(),
    reverseAggregateColumnEdges: new Map(),
    reverseExactLookupColumnEdges: new Map(),
    reverseSortedLookupColumnEdges: new Map(),
    ...overrides.reverseState,
  }
  const defaults: MaintenanceServiceArgs = {
    state,
    edgeArena: new EdgeArena(),
    reverseState,
    pivotOutputOwners: new Map(),
    captureSheetCellState: () => [],
    captureRowRangeCellState: () => [],
    captureColumnRangeCellState: () => [],
    setWasmProgramSyncPending: () => undefined,
    setMaterializedCellCount: () => undefined,
    resetFormulaRuntimeCaches: () => undefined,
    scheduleWasmProgramSync: () => undefined,
    resetWasmState: () => undefined,
  }
  return {
    ...defaults,
    ...overrides,
    state,
    reverseState,
  }
}

describe('EngineMaintenanceService', () => {
  it('estimates potential new cells only for materializing ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'maintenance-estimate' })
    await engine.ready()

    const estimate = Effect.runSync(
      getMaintenanceService(engine).estimatePotentialNewCells([
        { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
        { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B1', formula: 'A1+1' },
        { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'C1', format: '0.00' },
        { kind: 'clearCell', sheetName: 'Sheet1', address: 'D1' },
      ] satisfies EngineOp[]),
    )

    expect(estimate).toBe(3)
  })

  it('rewrites defined names and resets workbook through the extracted maintenance boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'maintenance-reset' })
    await engine.ready()
    engine.createSheet('Source')
    engine.setCellValue('Source', 'A1', 10)
    engine.setDefinedName('SourceCell', '=Source!A1')
    engine.setSelection('Source', 'B2')

    const maintenance = getMaintenanceService(engine)
    Effect.runSync(maintenance.rewriteDefinedNamesForSheetRename('Source', 'Renamed'))

    expect(engine.getDefinedName('SourceCell')).toEqual({
      name: 'SourceCell',
      value: '=Renamed!A1',
    })

    const previousBatchId = engine.getLastMetrics().batchId
    Effect.runSync(maintenance.resetWorkbook('Reset'))

    expect(engine.workbook.workbookName).toBe('Reset')
    expect(engine.getDefinedNames()).toEqual([])
    expect(engine.getSelectionState()).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'A1' },
      editMode: 'idle',
    })
    expect(engine.getLastMetrics()).toMatchObject({
      batchId: previousBatchId,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    })
  })

  it('captures sheet and range cell state through the extracted maintenance boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'maintenance-capture' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)

    const maintenance = getMaintenanceService(engine)

    expect(Effect.runSync(maintenance.captureSheetCellState('Sheet1')).map((op) => op.kind)).toEqual([
      'setCellValue',
      'setCellValue',
      'setCellValue',
    ])
    expect(Effect.runSync(maintenance.captureRowRangeCellState('Sheet1', 0, 2)).map((op) => op.kind)).toEqual([
      'setCellValue',
      'setCellValue',
    ])
    expect(Effect.runSync(maintenance.captureColumnRangeCellState('Sheet1', 0, 2)).map((op) => op.kind)).toEqual([
      'setCellValue',
      'setCellValue',
    ])
  })

  it('wraps capture callback failures with maintenance service errors', () => {
    const service = createEngineMaintenanceService(
      createMaintenanceServiceArgs({
        captureSheetCellState: () => {
          throw new Error('sheet capture boom')
        },
        captureRowRangeCellState: () => {
          throw new Error('row capture boom')
        },
        captureColumnRangeCellState: () => {
          throw new Error('column capture boom')
        },
      }),
    )

    expect(() => Effect.runSync(service.captureSheetCellState('Sheet1'))).toThrow('sheet capture boom')
    expect(() => Effect.runSync(service.captureRowRangeCellState('Sheet1', 0, 1))).toThrow('row capture boom')
    expect(() => Effect.runSync(service.captureColumnRangeCellState('Sheet1', 0, 1))).toThrow('column capture boom')
  })

  it('wraps rename, estimate, and reset failures with maintenance service errors', () => {
    const service = createEngineMaintenanceService(
      createMaintenanceServiceArgs({
        state: {
          workbook: new ThrowingMaintenanceWorkbook('throwing-maintenance-workbook'),
        },
      }),
    )
    const poisonedOps = new Proxy<EngineOp[]>([], {
      get(target, property, receiver) {
        if (property === 'length') {
          throw new Error('estimate boom')
        }
        return Reflect.get(target, property, receiver)
      },
    })

    expect(() => Effect.runSync(service.rewriteDefinedNamesForSheetRename('Sheet1', 'Renamed'))).toThrow('rename boom')
    expect(() => Effect.runSync(service.estimatePotentialNewCells(poisonedOps))).toThrow('estimate boom')
    expect(() => Effect.runSync(service.resetWorkbook('Reset'))).toThrow('reset boom')
  })
})
