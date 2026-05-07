import { readRuntimeImage, readRuntimeSnapshot, type EngineFormulaSourceRef, type SpreadsheetEngine } from '@bilig/core'
import { loadInitialLiteralSheet, prepareInitialMixedSheetLoad } from './initial-sheet-load.js'
import {
  inspectRuntimeSnapshotSheetDimensionsWithinLimits,
  inspectSheetWithinLimits,
  runtimeSnapshotMatchesSheetEntries,
  type WorkPaperSheetInspection,
} from './work-paper-sheet-inspection.js'
import type { SerializedWorkPaperNamedExpression, WorkPaperConfig, WorkPaperSheetDimensions, WorkPaperSheets } from './work-paper-types.js'

export function initializeWorkPaperFromSheets(args: {
  readonly engine: SpreadsheetEngine
  readonly config: WorkPaperConfig
  readonly sheets: WorkPaperSheets
  readonly namedExpressions: readonly SerializedWorkPaperNamedExpression[]
  readonly hasNamedExpressions: () => boolean
  readonly hasFunctionAliases: () => boolean
  readonly withEngineEventCaptureDisabled: (callback: () => void) => void
  readonly upsertNamedExpression: (expression: SerializedWorkPaperNamedExpression, options: { duringInitialization: boolean }) => void
  readonly rewriteFormulaForStorage: (formula: string, sheetId: number) => string
  readonly requireSheetId: (name: string) => number
  readonly cacheInitializedSheetDimensions: (sheetId: number, dimensions: WorkPaperSheetDimensions) => void
  readonly clearHistoryStacks: () => void
  readonly resetChangeTrackingCaches: () => void
}): void {
  const sheetEntries = Object.entries(args.sheets)
  const runtimeSnapshot = args.namedExpressions.length === 0 && !args.hasFunctionAliases() ? readRuntimeSnapshot(args.sheets) : undefined
  const runtimeSnapshotMatchesSheets = runtimeSnapshot !== undefined && runtimeSnapshotMatchesSheetEntries(sheetEntries, runtimeSnapshot)
  const runtimeSnapshotSheetsByName = runtimeSnapshotMatchesSheets
    ? new Map(runtimeSnapshot.sheets.map((sheet) => [sheet.name, sheet] as const))
    : undefined
  const runtimeImageSheetCellsByName =
    runtimeSnapshotMatchesSheets && runtimeSnapshot
      ? new Map((readRuntimeImage(runtimeSnapshot)?.sheetCells ?? []).map((sheet) => [sheet.sheetName, sheet] as const))
      : undefined
  const inspectedSheets = inspectWorkPaperInitialSheets({
    sheetEntries,
    config: args.config,
    runtimeSnapshotSheetsByName,
    runtimeImageSheetCellsByName,
  })

  args.withEngineEventCaptureDisabled(() => {
    if (runtimeSnapshot && runtimeSnapshotMatchesSheets) {
      args.engine.importSnapshot(runtimeSnapshot)
    } else {
      const sheetIds = sheetEntries.map(([sheetName]) => args.engine.createSheetForInitialization(sheetName))
      args.namedExpressions.forEach((expression) => {
        args.upsertNamedExpression(expression, { duringInitialization: true })
      })
      const initialFormulaRefs: EngineFormulaSourceRef[] = []
      let initialFormulaPotentialNewCells = 0
      for (let index = 0; index < sheetEntries.length; index += 1) {
        const [, sheet] = sheetEntries[index]!
        const sheetId = sheetIds[index]!
        const inspected = inspectedSheets[index]!
        if (!inspected.hasFormula) {
          loadInitialLiteralSheet(args.engine, sheetId, sheet, inspected)
          continue
        }
        const rewriteInitialFormula =
          !args.hasNamedExpressions() && !args.hasFunctionAliases()
            ? (formula: string) => formula
            : (formula: string) => args.rewriteFormulaForStorage(formula, sheetId)
        const prepared = prepareInitialMixedSheetLoad({
          engine: args.engine,
          sheetId,
          content: sheet,
          rewriteFormula: rewriteInitialFormula,
          inspection: inspected,
        })
        initialFormulaRefs.push(...prepared.formulaRefs)
        initialFormulaPotentialNewCells += prepared.potentialNewCells
      }
      if (initialFormulaRefs.length > 0) {
        args.engine.initializeFormulaSourcesAtNow(initialFormulaRefs, initialFormulaPotentialNewCells)
      }
    }
    for (let index = 0; index < sheetEntries.length; index += 1) {
      const sheetId = args.requireSheetId(sheetEntries[index]![0])
      const inspected = inspectedSheets[index]
      if (inspected !== undefined) {
        args.cacheInitializedSheetDimensions(sheetId, inspected.dimensions)
      }
    }
  })
  args.clearHistoryStacks()
  args.resetChangeTrackingCaches()
}

function inspectWorkPaperInitialSheets(args: {
  readonly sheetEntries: readonly (readonly [string, WorkPaperSheets[string]])[]
  readonly config: WorkPaperConfig
  readonly runtimeSnapshotSheetsByName:
    | ReadonlyMap<string, NonNullable<ReturnType<typeof readRuntimeSnapshot>>['sheets'][number]>
    | undefined
  readonly runtimeImageSheetCellsByName:
    | ReadonlyMap<string, NonNullable<NonNullable<ReturnType<typeof readRuntimeImage>>['sheetCells']>[number]>
    | undefined
}): WorkPaperSheetInspection[] {
  const inspectedSheets: WorkPaperSheetInspection[] = []
  for (let index = 0; index < args.sheetEntries.length; index += 1) {
    const [sheetName, sheet] = args.sheetEntries[index]!
    const snapshotSheet = args.runtimeSnapshotSheetsByName?.get(sheetName)
    inspectedSheets[index] = snapshotSheet
      ? (() => {
          const runtimeSheetCells = args.runtimeImageSheetCellsByName?.get(sheetName)
          const dimensions = inspectRuntimeSnapshotSheetDimensionsWithinLimits({
            sheetName,
            snapshotSheet,
            ...(runtimeSheetCells !== undefined ? { runtimeSheetCells } : {}),
            config: args.config,
          })
          return {
            hasFormula: false,
            dimensions,
            materializedCellCount: runtimeSheetCells?.cellCount ?? 0,
            maxColumnCount: dimensions.width,
            formulaCellCount: 0,
          }
        })()
      : inspectSheetWithinLimits(sheetName, sheet, args.config)
  }
  return inspectedSheets
}
