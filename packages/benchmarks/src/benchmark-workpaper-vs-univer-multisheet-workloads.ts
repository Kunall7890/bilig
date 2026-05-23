import type { WorkPaper, WorkPaperSheet } from '../../headless/src/work-paper.js'
import { normalizeWorkPaperValue } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
import type {
  UniverRuntime,
  UniverWorksheetFacade,
  WorkPaperUniverScenario,
  WorkPaperUniverWorkload,
} from './benchmark-workpaper-vs-univer.js'
import {
  buildCrossSheetAggregateSheets,
  buildCrossSheetDashboardSheets,
  buildCrossSheetScalarFanoutSheets,
  buildMultiSheetLiteralSheets,
  columnLabel,
} from './workpaper-benchmark-fixtures.js'

interface ObservedWorkbookCell {
  readonly col: number
  readonly key: string
  readonly row: number
  readonly sheetName: string
}

export function manySheetsBuildScenario(
  workload: WorkPaperUniverWorkload,
  sheetCount: number,
  rowsPerSheet: number,
  columnCount: number,
): WorkPaperUniverScenario {
  const sheets = buildMultiSheetLiteralSheets(sheetCount, rowsPerSheet, columnCount)
  return canonicalWorkbookBuildScenario({
    family: 'build',
    observedCells: [
      {
        col: columnCount - 1,
        key: 'terminalValue',
        row: rowsPerSheet - 1,
        sheetName: `Sheet${String(sheetCount)}`,
      },
    ],
    rowCount: rowsPerSheet * sheetCount,
    sheets,
    workload,
  })
}

export function crossSheetDashboardBuildScenario(
  workload: WorkPaperUniverWorkload,
  sheetCount: number,
  rowsPerSheet: number,
): WorkPaperUniverScenario {
  const sheets = buildCrossSheetDashboardSheets(sheetCount, rowsPerSheet)
  return canonicalWorkbookBuildScenario({
    family: 'cross-sheet',
    observedCells: [
      { col: 1, key: 'leadingDataTotal', row: 0, sheetName: 'Summary' },
      { col: 2, key: 'terminalDataTotal', row: sheetCount - 1, sheetName: 'Summary' },
    ],
    rowCount: rowsPerSheet * sheetCount + sheetCount,
    sheets,
    workload,
  })
}

export function crossSheetScalarFanoutRecalcScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const sheets = buildCrossSheetScalarFanoutSheets(rowCount)
  return canonicalWorkbookMutationScenario({
    edit: { col: 0, row: 0, sheetName: 'Data', value: 99 },
    family: 'cross-sheet',
    observedCells: [
      { col: 0, key: 'leadingValue', row: 0, sheetName: 'Summary' },
      { col: 0, key: 'terminalValue', row: rowCount - 1, sheetName: 'Summary' },
    ],
    rowCount: rowCount * 2,
    sheets,
    workload,
  })
}

export function crossSheetAggregateRecalcScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const sheets = buildCrossSheetAggregateSheets(rowCount)
  return canonicalWorkbookMutationScenario({
    edit: { col: 0, row: 0, sheetName: 'Data', value: 99 },
    family: 'cross-sheet',
    observedCells: [
      { col: 0, key: 'leadingSum', row: 0, sheetName: 'Summary' },
      { col: 0, key: 'terminalSum', row: rowCount - 1, sheetName: 'Summary' },
    ],
    rowCount: rowCount * 2,
    sheets,
    workload,
  })
}

export function crossSheetDashboardRecalcScenario(
  workload: WorkPaperUniverWorkload,
  sheetCount: number,
  rowsPerSheet: number,
): WorkPaperUniverScenario {
  const sheets = buildCrossSheetDashboardSheets(sheetCount, rowsPerSheet)
  return canonicalWorkbookMutationScenario({
    edit: { col: 1, row: 0, sheetName: 'Data1', value: 999 },
    family: 'cross-sheet',
    observedCells: [
      { col: 1, key: 'leadingDataTotal', row: 0, sheetName: 'Summary' },
      { col: 2, key: 'terminalDataTotal', row: sheetCount - 1, sheetName: 'Summary' },
    ],
    rowCount: rowsPerSheet * sheetCount + sheetCount,
    sheets,
    workload,
  })
}

function canonicalWorkbookBuildScenario(args: {
  readonly family: WorkPaperUniverScenario['fixture']['family']
  readonly observedCells: readonly ObservedWorkbookCell[]
  readonly rowCount: number
  readonly sheets: Record<string, WorkPaperSheet>
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const firstObservedCell = args.observedCells[0]!
  const columnCount = Math.max(1, ...Object.values(args.sheets).flatMap((sheet) => sheet.map((row) => row.length)))
  const formula = firstFormula(args.sheets) ?? args.workload
  const fixture = {
    family: args.family,
    formula,
    result: {
      address: formatA1(firstObservedCell.row, firstObservedCell.col),
      col: firstObservedCell.col,
      row: firstObservedCell.row,
      sheetName: firstObservedCell.sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const
  return {
    kind: 'build',
    fixture,
    buildWorkPaperSheets: () => args.sheets,
    setupUniver: (runtime) => setupUniverWorkbookSheets(runtime, args.sheets),
    verifyUniver: (runtime) => observeUniver(runtime, args.observedCells),
    verifyWorkPaper: (workbook) => observeWorkPaper(workbook, args.observedCells),
  }
}

function canonicalWorkbookMutationScenario(args: {
  readonly edit: {
    readonly col: number
    readonly row: number
    readonly sheetName: string
    readonly value: number
  }
  readonly family: WorkPaperUniverScenario['fixture']['family']
  readonly observedCells: readonly ObservedWorkbookCell[]
  readonly rowCount: number
  readonly sheets: Record<string, WorkPaperSheet>
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const firstObservedCell = args.observedCells[0]!
  const columnCount = Math.max(1, ...Object.values(args.sheets).flatMap((sheet) => sheet.map((row) => row.length)))
  const formula = firstFormula(args.sheets) ?? args.workload
  const fixture = {
    edit: {
      address: formatA1(args.edit.row, args.edit.col),
      col: args.edit.col,
      row: args.edit.row,
      sheetName: args.edit.sheetName,
      value: args.edit.value,
    },
    family: args.family,
    formula,
    result: {
      address: formatA1(firstObservedCell.row, firstObservedCell.col),
      col: firstObservedCell.col,
      row: firstObservedCell.row,
      sheetName: firstObservedCell.sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const
  return {
    kind: 'mutation',
    fixture,
    buildWorkPaperSheets: () => args.sheets,
    setupUniver: (runtime) => setupUniverWorkbookSheets(runtime, args.sheets),
    verifyUniver: (runtime) => observeUniver(runtime, args.observedCells),
    verifyWorkPaper: (workbook) => observeWorkPaper(workbook, args.observedCells),
  }
}

async function setupUniverWorkbookSheets(runtime: UniverRuntime, sheets: Record<string, WorkPaperSheet>): Promise<void> {
  const formulaRuns: {
    readonly formulas: string[]
    readonly row: number
    readonly sheet: UniverWorksheetFacade
    readonly startCol: number
  }[] = []
  for (const [sheetName, sheet] of Object.entries(sheets)) {
    const rowCount = Math.max(1, sheet.length)
    const columnCount = Math.max(1, ...sheet.map((row) => row.length))
    const univerSheet = ensureUniverSheet(runtime, sheetName, rowCount, columnCount)
    const values = sheet.map((row) => row.map((value) => (typeof value === 'string' && value.startsWith('=') ? '' : (value ?? ''))))
    univerSheet.getRange(0, 0, rowCount, columnCount).setValues(values)
    collectFormulaRuns(univerSheet, sheet, formulaRuns)
  }
  if (formulaRuns.length === 0) {
    return
  }
  for (const run of formulaRuns) {
    run.sheet.getRange(run.row, run.startCol, 1, run.formulas.length).setFormulas([run.formulas])
  }
}

function ensureUniverSheet(runtime: UniverRuntime, sheetName: string, rows: number, columns: number): UniverWorksheetFacade {
  return runtime.workbook.getSheetByName(sheetName) ?? runtime.workbook.create(sheetName, rows, columns)
}

function collectFormulaRuns(
  sheet: UniverWorksheetFacade,
  sourceSheet: WorkPaperSheet,
  formulaRuns: { readonly formulas: string[]; readonly row: number; readonly sheet: UniverWorksheetFacade; readonly startCol: number }[],
): void {
  for (let row = 0; row < sourceSheet.length; row += 1) {
    const cells = sourceSheet[row] ?? []
    let col = 0
    while (col < cells.length) {
      const value = cells[col]
      if (typeof value !== 'string' || !value.startsWith('=')) {
        col += 1
        continue
      }
      const startCol = col
      const formulas: string[] = []
      while (col < cells.length) {
        const formula = cells[col]
        if (typeof formula !== 'string' || !formula.startsWith('=')) {
          break
        }
        formulas.push(formula)
        col += 1
      }
      formulaRuns.push({ sheet, formulas, row, startCol })
    }
  }
}

function observeWorkPaper(workbook: WorkPaper, observedCells: readonly ObservedWorkbookCell[]): Record<string, unknown> {
  return Object.fromEntries(
    observedCells.map((cell) => [
      cell.key,
      normalizeBenchmarkValue(
        normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId(cell.sheetName)!, row: cell.row, col: cell.col })),
      ),
    ]),
  )
}

function observeUniver(runtime: UniverRuntime, observedCells: readonly ObservedWorkbookCell[]): Record<string, unknown> {
  return Object.fromEntries(
    observedCells.map((cell) => [
      cell.key,
      normalizeBenchmarkValue(runtime.workbook.getSheetByName(cell.sheetName)!.getRange(formatA1(cell.row, cell.col)).getValue()),
    ]),
  )
}

function firstFormula(sheets: Record<string, WorkPaperSheet>): string | undefined {
  for (const sheet of Object.values(sheets)) {
    for (const row of sheet) {
      for (const value of row) {
        if (typeof value === 'string' && value.startsWith('=')) {
          return value
        }
      }
    }
  }
  return undefined
}

function formatA1(row: number, col: number): string {
  return `${columnLabel(col)}${String(row + 1)}`
}
