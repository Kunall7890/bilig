import type { WorkPaper, WorkPaperSheet } from '../../headless/src/work-paper.js'
import { normalizeWorkPaperValue } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
import type {
  UniverEditableValue,
  UniverRuntime,
  WorkPaperUniverFixture,
  WorkPaperUniverScenario,
  WorkPaperUniverWorkload,
  WorkPaperUniverWorkloadFamily,
} from './benchmark-workpaper-vs-univer.js'
import {
  build2dAggregateSheet,
  buildApproxLookupDuplicateSheet,
  buildApproxLookupSheet,
  buildDenseLiteralSheet,
  buildFormulaChainRow,
  buildFormulaFanoutRow,
  buildIndexMatchExactSheet,
  buildIndexReferenceSheet,
  buildLookupSheet,
  buildMixedContentSheet,
  buildOverlappingAggregateSheet,
  buildParserCacheMixedTemplateSheet,
  buildParserCacheTemplateSheet,
  buildParserCacheUniqueFormulaSheet,
  buildSlidingAggregateSheet,
  buildTextLookupSheet,
  columnLabel,
  textLookupKey,
} from './workpaper-benchmark-fixtures.js'

export function denseLiteralBuildScenario(workload: WorkPaperUniverWorkload, rows: number, cols: number): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: cols - 1, key: 'terminalValue', row: rows - 1 }],
    rowCount: rows,
    sheet: buildDenseLiteralSheet(rows, cols),
    workload,
  })
}

export function mixedContentBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: 5, key: 'terminalFormulaValue', row: rowCount - 1 }],
    rowCount,
    sheet: buildMixedContentSheet(rowCount),
    workload,
  })
}

export function parserCacheTemplateBuildScenario(
  workload: WorkPaperUniverWorkload,
  sheet: WorkPaperSheet,
  rowCount: number,
): WorkPaperUniverScenario {
  return canonicalSingleSheetBuildScenario({
    family: 'build',
    observedCells: [{ col: 5, key: 'terminalValue', row: rowCount - 1 }],
    rowCount,
    sheet,
    workload,
  })
}

export function parserCacheRowTemplateBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return parserCacheTemplateBuildScenario(workload, buildParserCacheTemplateSheet(rowCount), rowCount)
}

export function parserCacheMixedTemplateBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return parserCacheTemplateBuildScenario(workload, buildParserCacheMixedTemplateSheet(rowCount), rowCount)
}

export function parserCacheUniqueFormulaBuildScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return parserCacheTemplateBuildScenario(workload, buildParserCacheUniqueFormulaSheet(rowCount), rowCount)
}

export function formulaChainRowScenario(workload: WorkPaperUniverWorkload, downstreamCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'formula-chain',
    observedCells: [{ col: downstreamCount, key: 'terminalValue', row: 0 }],
    rowCount: 1,
    sheet: [buildFormulaChainRow(downstreamCount)],
    workload,
  })
}

export function formulaFanoutRowScenario(workload: WorkPaperUniverWorkload, fanoutCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'formula-fanout',
    observedCells: [
      { col: fanoutCount, key: 'terminalValue', row: 0 },
      { col: fanoutCount, key: 'width', row: 0, value: fanoutCount + 1 },
    ],
    rowCount: 1,
    sheet: [buildFormulaFanoutRow(fanoutCount)],
    workload,
  })
}

export function aggregate2dCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'aggregate-2d',
    observedCells: [
      { col: 2, key: 'terminalSum', row: rowCount - 1 },
      { col: 2, key: 'leadingSum', row: 0 },
    ],
    rowCount,
    sheet: build2dAggregateSheet(rowCount),
    workload,
  })
}

export function overlappingAggregateCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'overlapping-aggregate',
    observedCells: [{ col: 1, key: 'terminalSum', row: rowCount - 1 }],
    rowCount,
    sheet: buildOverlappingAggregateSheet(rowCount),
    workload,
  })
}

export function slidingAggregateCanonicalScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  window: number,
): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'overlapping-aggregate',
    observedCells: [
      { col: 1, key: 'terminalSum', row: rowCount - 1 },
      { col: 1, key: 'leadingSum', row: 0 },
    ],
    rowCount,
    sheet: buildSlidingAggregateSheet(rowCount, window),
    workload,
  })
}

export function exactLookupCanonicalScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  useColumnIndex: boolean,
): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: rowCount },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildLookupSheet(rowCount),
    workbookOptions: { useColumnIndex },
    workload,
  })
}

export function indexMatchExactCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: textLookupKey(rowCount - 1) },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildIndexMatchExactSheet(rowCount),
    workload,
  })
}

export function indexReferenceCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: rowCount - 1 },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildIndexReferenceSheet(rowCount),
    workload,
  })
}

export function approximateLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: rowCount - 0.5 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupSheet(rowCount),
    workload,
  })
}

export function approximateDuplicateLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: Math.floor(rowCount / 5) + 0.5 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupDuplicateSheet(rowCount),
    workload,
  })
}

export function textLookupCanonicalScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return canonicalSingleSheetScenario({
    edit: { col: 3, row: 0, value: textLookupKey(rowCount - 1) },
    family: 'lookup-exact',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildTextLookupSheet(rowCount),
    workload,
  })
}

function canonicalSingleSheetBuildScenario(args: {
  readonly family: WorkPaperUniverWorkloadFamily
  readonly observedCells: readonly {
    readonly col: number
    readonly key: string
    readonly row: number
    readonly value?: number
  }[]
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly workbookOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const sheetName = 'Bench'
  const columnCount = Math.max(...args.observedCells.map((cell) => cell.col + 1), ...args.sheet.map((row) => row.length))
  const formula = firstFormula(args.sheet) ?? args.workload
  const fixture = {
    family: args.family,
    formula,
    result: {
      address: formatA1(args.observedCells[0]!.row, args.observedCells[0]!.col),
      col: args.observedCells[0]!.col,
      row: args.observedCells[0]!.row,
      sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const satisfies WorkPaperUniverFixture
  const observeWorkPaper = (workbook: WorkPaper): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ??
          normalizeBenchmarkValue(
            normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId(sheetName)!, row: cell.row, col: cell.col })),
          ),
      ]),
    )
  const observeUniver = (runtime: UniverRuntime): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ?? normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(cell.row, cell.col)).getValue()),
      ]),
    )
  return {
    kind: 'build',
    fixture,
    buildWorkPaperSheets: () => ({ [sheetName]: args.sheet }),
    ...(args.workbookOptions ? { workpaperOptions: args.workbookOptions } : {}),
    setupUniver: (runtime) => setupUniverSheet(runtime, args.sheet),
    verifyUniver: observeUniver,
    verifyWorkPaper: observeWorkPaper,
  }
}

function canonicalSingleSheetScenario(args: {
  readonly edit: { readonly col: number; readonly row: number; readonly value: UniverEditableValue }
  readonly family: WorkPaperUniverWorkloadFamily
  readonly observedCells: readonly {
    readonly col: number
    readonly key: string
    readonly row: number
    readonly value?: number
  }[]
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly workbookOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const sheetName = 'Bench'
  const columnCount = Math.max(args.edit.col + 1, ...args.observedCells.map((cell) => cell.col + 1), ...args.sheet.map((row) => row.length))
  const formula = firstFormula(args.sheet) ?? args.workload
  const fixture = {
    edit: {
      address: formatA1(args.edit.row, args.edit.col),
      col: args.edit.col,
      row: args.edit.row,
      sheetName,
      value: args.edit.value,
    },
    family: args.family,
    formula,
    result: {
      address: formatA1(args.observedCells[0]!.row, args.observedCells[0]!.col),
      col: args.observedCells[0]!.col,
      row: args.observedCells[0]!.row,
      sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  } as const satisfies WorkPaperUniverFixture
  const observeWorkPaper = (workbook: WorkPaper): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ??
          normalizeBenchmarkValue(
            normalizeWorkPaperValue(workbook.getCellValue({ sheet: workbook.getSheetId(sheetName)!, row: cell.row, col: cell.col })),
          ),
      ]),
    )
  const observeUniver = (runtime: UniverRuntime): Record<string, unknown> =>
    Object.fromEntries(
      args.observedCells.map((cell) => [
        cell.key,
        cell.value ?? normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(cell.row, cell.col)).getValue()),
      ]),
    )
  return {
    kind: 'mutation',
    fixture,
    buildWorkPaperSheets: () => ({ [sheetName]: args.sheet }),
    ...(args.workbookOptions ? { workpaperOptions: args.workbookOptions } : {}),
    setupUniver: (runtime) => setupUniverSheet(runtime, args.sheet),
    verifyUniver: observeUniver,
    verifyWorkPaper: observeWorkPaper,
  }
}

async function setupUniverSheet(runtime: UniverRuntime, sheet: WorkPaperSheet): Promise<void> {
  const rowCount = sheet.length
  const columnCount = Math.max(1, ...sheet.map((row) => row.length))
  const formulaRuns: { readonly formulas: string[]; readonly row: number; readonly startCol: number }[] = []
  const values: UniverEditableValue[][] = Array.from({ length: rowCount }, (_rowValue, row) =>
    Array.from({ length: columnCount }, (_colValue, col) => {
      const value = sheet[row]?.[col] ?? ''
      return typeof value === 'string' && value.startsWith('=') ? '' : value
    }),
  )
  for (let row = 0; row < sheet.length; row += 1) {
    const cells = sheet[row] ?? []
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
      formulaRuns.push({ formulas, row, startCol })
    }
  }
  runtime.sheet.getRange(0, 0, rowCount, columnCount).setValues(values)

  if (formulaRuns.length === 0) {
    return
  }

  for (const run of formulaRuns) {
    runtime.sheet.getRange(run.row, run.startCol, 1, run.formulas.length).setFormulas([run.formulas])
  }
}

function firstFormula(sheet: WorkPaperSheet): string | undefined {
  for (const row of sheet) {
    for (const value of row) {
      if (typeof value === 'string' && value.startsWith('=')) {
        return value
      }
    }
  }
  return undefined
}

function formatA1(row: number, col: number): string {
  return `${columnLabel(col)}${String(row + 1)}`
}
