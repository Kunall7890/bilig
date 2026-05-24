import { WorkPaper, type WorkPaperSheet } from '../../headless/src/work-paper.js'
import { normalizeWorkPaperValue } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
import { setupUniverWorkbookSheets } from './benchmark-workpaper-vs-univer-multisheet-workloads.js'
import type { UniverRuntime, WorkPaperUniverScenario, WorkPaperUniverWorkload } from './benchmark-workpaper-vs-univer.js'
import {
  address,
  buildMixedContentSheet,
  buildParserCacheMixedTemplateSheet,
  buildParserCacheTemplateSheet,
  buildRenameDependencySheets,
  columnLabel,
} from './workpaper-benchmark-fixtures.js'

const namedExpressionName = 'TaxRate'

export function rebuildAndRecalculateScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const sheet = buildParserCacheTemplateSheet(rowCount)
  return {
    kind: 'mutation',
    fixture: {
      family: 'rebuild',
      formula: firstFormula({ Bench: sheet }) ?? workload,
      result: {
        address: formatA1(rowCount - 1, 4),
        col: 4,
        row: rowCount - 1,
        sheetName: 'Bench',
      },
      columnCount: 6,
      rowCount,
    },
    buildWorkPaperSheets: () => ({ Bench: sheet }),
    executeUniverMutation: (runtime) => runtime.formula.executeCalculation(),
    executeWorkPaperMutation: (workbook) => workbook.rebuildAndRecalculate(),
    setupUniver: (runtime) => setupUniverWorkbookSheets(runtime, { Bench: sheet }),
    verifyUniver: (runtime) => ({
      terminalValue: normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(rowCount - 1, 4)).getValue()),
      dimensions: { height: runtime.sheet.getMaxRows(), width: runtime.sheet.getMaxColumns() },
    }),
    verifyWorkPaper: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      const dimensions = workbook.getSheetDimensions(sheetId)
      return {
        terminalValue: normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4)))),
        dimensions: { height: dimensions.height, width: dimensions.width },
      }
    },
  }
}

export function runtimeSnapshotBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const serializedSheets = buildSerializedRuntimeSheets(rowCount)
  return {
    kind: 'build',
    fixture: {
      family: 'rebuild',
      formula: firstFormula(serializedSheets) ?? workload,
      result: {
        address: formatA1(rowCount - 1, 5),
        col: 5,
        row: rowCount - 1,
        sheetName: 'Bench',
      },
      columnCount: 6,
      rowCount,
    },
    buildWorkPaperSheets: () => serializedSheets,
    setupUniver: (runtime) => setupUniverWorkbookSheets(runtime, serializedSheets),
    verifyUniver: (runtime) => ({
      sheetCount: runtime.workbook.getNumSheets(),
      benchTerminal: normalizeBenchmarkValue(runtime.workbook.getSheetByName('Bench')!.getRange(formatA1(rowCount - 1, 5)).getValue()),
    }),
    verifyWorkPaper: (workbook) => {
      const benchId = workbook.getSheetId('Bench')!
      return {
        sheetCount: workbook.countSheets(),
        benchTerminal: normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(benchId, rowCount - 1, 5)))),
      }
    },
  }
}

export function sheetRenameDependencyScenario(workload: WorkPaperUniverWorkload): WorkPaperUniverScenario {
  const sheets = buildRenameDependencySheets()
  return {
    kind: 'mutation',
    fixture: {
      family: 'sheet-lifecycle',
      formula: firstFormula(sheets) ?? workload,
      result: {
        address: 'A1',
        col: 0,
        row: 0,
        sheetName: 'Summary',
      },
      columnCount: 2,
      rowCount: 4,
    },
    buildWorkPaperSheets: () => sheets,
    executeUniverMutation: (runtime) => runtime.workbook.getSheetByName('Data')!.setName('Source'),
    executeWorkPaperMutation: (workbook) => workbook.renameSheet(workbook.getSheetId('Data')!, 'Source'),
    setupUniver: (runtime) => setupUniverWorkbookSheets(runtime, sheets),
    verifyUniver: (runtime) => {
      const summary = runtime.workbook.getSheetByName('Summary')!
      return {
        sheetNames: runtime.workbook.getSheetByName('Source') === null ? [] : ['Source', 'Summary'],
        scalarFormula: summary.getRange(0, 0).getFormula(),
        aggregateFormula: summary.getRange(0, 1).getFormula(),
        scalarValue: normalizeBenchmarkValue(summary.getRange(0, 0).getValue()),
        aggregateValue: normalizeBenchmarkValue(summary.getRange(0, 1).getValue()),
      }
    },
    verifyWorkPaper: (workbook) => {
      const summarySheetId = workbook.getSheetId('Summary')!
      return {
        sheetNames: workbook.getSheetNames(),
        scalarFormula: workbook.getCellFormula(address(summarySheetId, 0, 0)),
        aggregateFormula: workbook.getCellFormula(address(summarySheetId, 0, 1)),
        scalarValue: normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 0)))),
        aggregateValue: normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 1)))),
      }
    },
  }
}

export function namedExpressionChangeScenario(_workload: WorkPaperUniverWorkload): WorkPaperUniverScenario {
  const sheet = buildNamedExpressionSheet()
  return {
    kind: 'mutation',
    fixture: {
      family: 'named-expression',
      formula: `=${namedExpressionName}`,
      result: {
        address: 'B2',
        col: 1,
        row: 1,
        sheetName: 'Bench',
      },
      columnCount: 3,
      rowCount: 2,
    },
    buildWorkPaperSheets: () => ({ Bench: sheet }),
    executeUniverMutation: (runtime) => updateUniverDefinedFormula(runtime, namedExpressionName, '3'),
    executeWorkPaperMutation: (workbook) => workbook.changeNamedExpression(namedExpressionName, '=3'),
    setupUniver: async (runtime) => {
      insertUniverDefinedFormula(runtime, namedExpressionName, '2')
      await setupUniverWorkbookSheets(runtime, { Bench: sheet })
    },
    verifyUniver: (runtime) => ({
      rateValue: normalizeBenchmarkValue(runtime.sheet.getRange(1, 1).getValue()),
    }),
    verifyWorkPaper: (workbook) => ({
      rateValue: normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(workbook.getSheetId('Bench')!, 1, 1)))),
    }),
    workpaperNamedExpressions: [{ name: namedExpressionName, expression: '=2' }],
  }
}

function buildSerializedRuntimeSheets(rowCount: number): Record<string, WorkPaperSheet> {
  const seeded = WorkPaper.buildFromSheets({
    Bench: buildMixedContentSheet(rowCount),
    Templates: buildParserCacheMixedTemplateSheet(Math.max(Math.floor(rowCount / 2), 2)),
  })
  const serializedSheets = seeded.getAllSheetsSerialized()
  seeded.dispose()
  return serializedSheets
}

function buildNamedExpressionSheet(): WorkPaperSheet {
  return [[1, `=${namedExpressionName}+1`, `=${namedExpressionName}*2`], [2, `=${namedExpressionName}`]]
}

function insertUniverDefinedFormula(runtime: UniverRuntime, name: string, formulaWithoutEquals: string): void {
  runtime.workbook.insertDefinedNameBuilder(
    runtime.workbook.newDefinedNameBuilder().setName(name).setFormula(formulaWithoutEquals).setScopeToWorkbook().build(),
  )
}

function updateUniverDefinedFormula(runtime: UniverRuntime, name: string, formulaWithoutEquals: string): void {
  const existing = runtime.workbook.getDefinedName(name)
  const param = isDefinedNameFacade(existing) ? existing._definedNameParam : existing
  runtime.workbook.updateDefinedNameBuilder(
    runtime.workbook.newDefinedNameBuilder().load(param).setFormula(formulaWithoutEquals).build(),
  )
}

function isDefinedNameFacade(value: unknown): value is { readonly _definedNameParam: unknown } {
  return typeof value === 'object' && value !== null && '_definedNameParam' in value
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
