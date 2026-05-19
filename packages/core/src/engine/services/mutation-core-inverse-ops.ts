import type { EngineOp } from '@bilig/workbook-domain'
import { structuralTransformForOp } from '../../engine-structural-utils.js'
import { sheetMetadataToOps } from '../../engine-snapshot-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { buildMutationMetadataInverseOps } from './mutation-inverse-metadata-ops.js'
import { captureStructuralWorkbookMetadataOps, clearStructuralSheetMetadataOps } from './mutation-structural-metadata-ops.js'

export type MutationCoreCapturedInverseKind =
  | 'deleteSheet'
  | 'deleteRows'
  | 'deleteColumns'
  | 'setCellValue'
  | 'setCellFormula'
  | 'clearCell'

type MutationCoreCapturedInverseOp = Extract<EngineOp, { kind: MutationCoreCapturedInverseKind }>

type StructuralUndoAxis = 'row' | 'column'

const mutationCoreCapturedInverseKinds: ReadonlySet<EngineOp['kind']> = new Set([
  'deleteSheet',
  'deleteRows',
  'deleteColumns',
  'setCellValue',
  'setCellFormula',
  'clearCell',
])

function isMutationCoreCapturedInverseOp(op: EngineOp): op is MutationCoreCapturedInverseOp {
  return mutationCoreCapturedInverseKinds.has(op.kind)
}

function captureDeletedSheetInverseOps(
  workbook: WorkbookStore,
  sheetName: string,
  captureSheetCellState: (sheetName: string) => EngineOp[],
): EngineOp[] {
  const sheet = workbook.getSheet(sheetName)
  if (!sheet) {
    return []
  }
  const restoredOps: EngineOp[] = [{ kind: 'upsertSheet', name: sheet.name, order: sheet.order }]
  restoredOps.push(...sheetMetadataToOps(workbook, sheet.name))
  workbook
    .listTables()
    .filter((table) => table.sheetName === sheet.name)
    .forEach((table) => {
      restoredOps.push({
        kind: 'upsertTable',
        table: structuredClone(table),
      })
    })
  workbook
    .listSpills()
    .filter((spill) => spill.sheetName === sheet.name)
    .forEach((spill) => {
      restoredOps.push({
        kind: 'upsertSpillRange',
        sheetName: spill.sheetName,
        address: spill.address,
        rows: spill.rows,
        cols: spill.cols,
      })
    })
  workbook
    .listPivots()
    .filter((pivot) => pivot.sheetName === sheet.name && pivot.source)
    .forEach((pivot) => {
      if (!pivot.source) {
        return
      }
      restoredOps.push({
        kind: 'upsertPivotTable',
        name: pivot.name,
        sheetName: pivot.sheetName,
        address: pivot.address,
        source: { ...pivot.source },
        groupBy: [...pivot.groupBy],
        values: pivot.values.map((value) => Object.assign({}, value)),
        rows: pivot.rows,
        cols: pivot.cols,
      })
    })
  workbook
    .listCharts()
    .filter((chart) => chart.sheetName === sheet.name || chart.source.sheetName === sheet.name)
    .forEach((chart) => {
      restoredOps.push({
        kind: 'upsertChart',
        chart: structuredClone(chart),
      })
    })
  workbook
    .listImages()
    .filter((image) => image.sheetName === sheet.name)
    .forEach((image) => {
      restoredOps.push({
        kind: 'upsertImage',
        image: structuredClone(image),
      })
    })
  workbook
    .listShapes()
    .filter((shape) => shape.sheetName === sheet.name)
    .forEach((shape) => {
      restoredOps.push({
        kind: 'upsertShape',
        shape: structuredClone(shape),
      })
    })
  restoredOps.push(...captureSheetCellState(sheet.name))
  return restoredOps
}

export function createMutationCoreInverseOps(args: {
  readonly workbook: WorkbookStore
  readonly captureSheetCellState: (sheetName: string) => EngineOp[]
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly restoreCellOps: (sheetName: string, address: string) => EngineOp[]
  readonly captureFormulaCellStateForStructuralUndo: (
    sheetName: string,
    axis: StructuralUndoAxis,
    start: number,
    count: number,
  ) => EngineOp[]
  readonly captureFormulaCellStateForStructuralMoveUndo: (
    sheetName: string,
    axis: StructuralUndoAxis,
    start: number,
    count: number,
    target: number,
  ) => EngineOp[]
}): {
  readonly inverseOpsFor: (op: EngineOp) => EngineOp[]
  readonly buildInverseOps: (ops: readonly EngineOp[]) => EngineOp[]
} {
  const inverseOpsFor = (op: EngineOp): EngineOp[] => {
    if (op.kind === 'insertRows' || op.kind === 'insertColumns') {
      const metadataInverseOps = buildMutationMetadataInverseOps(args.workbook, op) ?? []
      const axis: StructuralUndoAxis = op.kind === 'insertRows' ? 'row' : 'column'
      return [...metadataInverseOps, ...args.captureFormulaCellStateForStructuralMoveUndo(op.sheetName, axis, op.start, op.count, op.start)]
    }
    if (op.kind === 'moveRows' || op.kind === 'moveColumns') {
      const metadataInverseOps = buildMutationMetadataInverseOps(args.workbook, op) ?? []
      const axis: StructuralUndoAxis = op.kind === 'moveRows' ? 'row' : 'column'
      return [
        ...metadataInverseOps,
        ...args.captureFormulaCellStateForStructuralMoveUndo(op.sheetName, axis, op.start, op.count, op.target),
      ]
    }
    const metadataInverseOps = buildMutationMetadataInverseOps(args.workbook, op)
    if (metadataInverseOps !== undefined) {
      return metadataInverseOps
    }
    if (!isMutationCoreCapturedInverseOp(op)) {
      throw new Error(`Unhandled inverse operation: ${op.kind}`)
    }

    switch (op.kind) {
      case 'deleteSheet':
        return captureDeletedSheetInverseOps(args.workbook, op.name, args.captureSheetCellState)
      case 'deleteRows': {
        const entries = args.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(args.workbook, op.sheetName, transform),
          {
            kind: 'insertRows',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...sheetMetadataToOps(args.workbook, op.sheetName, { includeAxisEntries: false }),
          ...args.captureRowRangeCellState(op.sheetName, op.start, op.count),
          ...args.captureFormulaCellStateForStructuralUndo(op.sheetName, 'row', op.start, op.count),
          ...captureStructuralWorkbookMetadataOps(args.workbook),
        ]
      }
      case 'deleteColumns': {
        const entries = args.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(args.workbook, op.sheetName, transform),
          {
            kind: 'insertColumns',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...sheetMetadataToOps(args.workbook, op.sheetName, { includeAxisEntries: false }),
          ...args.captureColumnRangeCellState(op.sheetName, op.start, op.count),
          ...args.captureFormulaCellStateForStructuralUndo(op.sheetName, 'column', op.start, op.count),
          ...captureStructuralWorkbookMetadataOps(args.workbook),
        ]
      }
      case 'setCellValue':
      case 'setCellFormula':
      case 'clearCell':
        return args.restoreCellOps(op.sheetName, op.address)
      default: {
        const exhaustive: never = op
        return exhaustive
      }
    }
  }

  const buildInverseOps = (ops: readonly EngineOp[]): EngineOp[] => {
    const inverseOps: EngineOp[] = []
    for (let index = ops.length - 1; index >= 0; index -= 1) {
      const op = ops[index]
      if (op !== undefined) {
        inverseOps.push(...inverseOpsFor(op))
      }
    }
    return inverseOps
  }

  return { inverseOpsFor, buildInverseOps }
}
