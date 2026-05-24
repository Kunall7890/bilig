import { performance } from 'node:perf_hooks'
import type { WorkPaper, WorkPaperSheet } from '../../headless/src/work-paper.js'
import { normalizeWorkPaperValue } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { normalizeBenchmarkValue } from './benchmark-workpaper-vs-univer-fixtures.js'
import type {
  UniverEditableValue,
  UniverRuntime,
  WorkPaperUniverScenario,
  WorkPaperUniverWorkload,
  WorkPaperUniverWorkloadFamily,
} from './benchmark-workpaper-vs-univer.js'
import {
  address,
  buildApproxLookupDescendingSheet,
  buildApproxLookupSheet,
  buildBatchMultiColumnRows,
  buildConditionalAggregationMixedSheet,
  buildConditionalAggregationSharedCriteriaSheet,
  buildConditionalAggregationSheet,
  buildFormulaEditChainRow,
  buildFormulaGridSheet,
  buildFormulaChainRow,
  buildLookupSheet,
  buildMixedFrontierSheet,
  buildRectangularBlockFormulaSheet,
  buildSparseWideSheet,
  buildValueFormulaRows,
  columnLabel,
  range,
} from './workpaper-benchmark-fixtures.js'

interface ObservedCell {
  readonly col: number
  readonly key: string
  readonly row: number
}

export function singleEditRecalcScenario(workload: WorkPaperUniverWorkload, downstreamCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'formula-chain',
    observedCells: [{ col: downstreamCount, key: 'terminalValue', row: 0 }],
    rowCount: 1,
    sheet: [buildFormulaChainRow(downstreamCount)],
    workload,
  })
}

export function mixedFrontierScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 0, row: 0, value: 99 },
    family: 'dirty-execution',
    observedCells: [
      { col: 2, key: 'terminalAggregate', row: rowCount - 1 },
      { col: 1, key: 'terminalFanout', row: rowCount - 1 },
    ],
    rowCount,
    sheet: buildMixedFrontierSheet(rowCount),
    workload,
  })
}

export function formulaEditScenario(workload: WorkPaperUniverWorkload, downstreamCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 2, row: 0, value: '=A1*B1' },
    executeUniverMutation: (runtime) => runtime.sheet.getRange(0, 2).setFormula('=A1*B1'),
    executeWorkPaperMutation: (workbook) => workbook.setCellContents(address(workbook.getSheetId('Bench')!, 0, 2), '=A1*B1'),
    family: 'dirty-execution',
    observedCells: [{ col: downstreamCount + 2, key: 'terminalValue', row: 0 }],
    rowCount: 1,
    sheet: [buildFormulaEditChainRow(downstreamCount)],
    workload,
  })
}

export function batchSingleColumnScenario(workload: WorkPaperUniverWorkload, editCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      for (let row = 0; row < editCount; row += 1) {
        runtime.sheet.getRange(row, 0).setValue(row * 3)
      }
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      return workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      })
    },
    family: 'batch-edit',
    observedCells: [
      { col: 1, key: 'sampleFormulaValue', row: editCount - 1 },
      { col: 1, key: 'width', row: editCount - 1, value: 2 },
    ],
    rowCount: editCount,
    sheet: buildValueFormulaRows(editCount),
    workload,
  })
}

export function batchMultiColumnScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      for (let row = 0; row < rowCount; row += 1) {
        runtime.sheet.getRange(row, 0, 1, 2).setValues([[row * 3, row * 5]])
      }
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      return workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
          workbook.setCellContents(address(sheetId, row, 1), row * 5)
        }
      })
    },
    family: 'batch-edit',
    observedCells: [
      { col: 2, key: 'sampleSumValue', row: rowCount - 1 },
      { col: 3, key: 'sampleProductValue', row: rowCount - 1 },
    ],
    rowCount,
    sheet: buildBatchMultiColumnRows(rowCount),
    workload,
  })
}

export function rectangularBatchEditScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  inputCols: number,
): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => runtime.sheet.getRange(0, 0, rowCount, inputCols).setValues(rectangularEditValues(rowCount, inputCols)),
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      return workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), (row + 1) * (col + 2))
          }
        }
      })
    },
    family: 'batch-edit',
    observedCells: [
      { col: inputCols, key: 'leadingSum', row: 0 },
      { col: inputCols, key: 'terminalSum', row: rowCount - 1 },
      { col: inputCols, key: 'width', row: rowCount - 1, value: inputCols + 1 },
    ],
    rowCount,
    sheet: buildRectangularBlockFormulaSheet(rowCount, inputCols),
    workload,
  })
}

export function rectangularBatchClearScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  inputCols: number,
): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => runtime.sheet.getRange(0, 0, rowCount, inputCols).setValues(emptyGrid(rowCount, inputCols)),
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      return workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), null)
          }
        }
      })
    },
    family: 'batch-edit',
    observedCells: [
      { col: inputCols, key: 'leadingSum', row: 0 },
      { col: inputCols, key: 'terminalSum', row: rowCount - 1 },
      { col: inputCols, key: 'width', row: rowCount - 1, value: inputCols + 1 },
    ],
    rowCount,
    sheet: buildRectangularBlockFormulaSheet(rowCount, inputCols),
    workload,
  })
}

export function batchSingleColumnUndoScenario(workload: WorkPaperUniverWorkload, editCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      runtime.sheet.getRange(0, 0, editCount, 1).setValues(Array.from({ length: editCount }, (_value, row) => [row * 3]))
      runtime.workbook.undo()
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      })
      return workbook.undo()
    },
    family: 'batch-edit',
    observedCells: [
      { col: 0, key: 'restoredValue', row: 0 },
      { col: 1, key: 'restoredFormulaValue', row: editCount - 1 },
    ],
    rowCount: editCount,
    sheet: buildValueFormulaRows(editCount),
    workload,
  })
}

export function suspendedBatchSingleColumnScenario(workload: WorkPaperUniverWorkload, editCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      for (let row = 0; row < editCount; row += 1) {
        runtime.sheet.getRange(row, 0).setValue(row * 7)
      }
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      workbook.suspendEvaluation()
      for (let row = 0; row < editCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 7)
      }
      return workbook.resumeEvaluation()
    },
    family: 'batch-edit',
    observedCells: [
      { col: 1, key: 'sampleFormulaValue', row: editCount - 1 },
      { col: 1, key: 'width', row: editCount - 1, value: 2 },
    ],
    rowCount: editCount,
    sheet: buildValueFormulaRows(editCount),
    workload,
  })
}

export function suspendedBatchMultiColumnScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      for (let row = 0; row < rowCount; row += 1) {
        runtime.sheet.getRange(row, 0, 1, 2).setValues([[row * 3, row * 5]])
      }
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      workbook.suspendEvaluation()
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 3)
        workbook.setCellContents(address(sheetId, row, 1), row * 5)
      }
      return workbook.resumeEvaluation()
    },
    family: 'batch-edit',
    observedCells: [
      { col: 2, key: 'sampleSumValue', row: rowCount - 1 },
      { col: 3, key: 'sampleProductValue', row: rowCount - 1 },
    ],
    rowCount,
    sheet: buildBatchMultiColumnRows(rowCount),
    workload,
  })
}

export function rangeReadDenseScenario(workload: WorkPaperUniverWorkload, rows: number, cols: number): WorkPaperUniverScenario {
  return rangeReadScenario({
    executeUniverRead: (runtime) => runtime.sheet.getRange(0, 0, rows, cols).getValues(),
    executeWorkPaperRead: (workbook, sheetId) => workbook.getRangeValues(range(sheetId, 0, 0, rows - 1, cols - 1)),
    family: 'range-read',
    rowCount: rows,
    sheet: Array.from({ length: rows }, (_rowValue, row) => Array.from({ length: cols }, (_colValue, col) => row * cols + col + 1)),
    summarizeUniverRead: (values) => summarizeDenseRangeRead(values, normalizeBenchmarkValue),
    summarizeWorkPaperRead: (values) => summarizeDenseRangeRead(values, normalizeWorkPaperValue),
    workload,
  })
}

export function rangeReadSparseWideScenario(workload: WorkPaperUniverWorkload, rowCount: number, colCount: number): WorkPaperUniverScenario {
  const middleCol = Math.floor(colCount / 2)
  return rangeReadScenario({
    executeUniverRead: (runtime) => runtime.sheet.getRange(0, 0, rowCount, colCount).getValues(),
    executeWorkPaperRead: (workbook, sheetId) => workbook.getRangeValues(range(sheetId, 0, 0, rowCount - 1, colCount - 1)),
    family: 'range-read',
    rowCount,
    sheet: buildSparseWideSheet(rowCount, colCount),
    summarizeUniverRead: (values) => summarizeSparseWideRead(values, middleCol, normalizeBenchmarkValue),
    summarizeWorkPaperRead: (values) => summarizeSparseWideRead(values, middleCol, normalizeWorkPaperValue),
    workload,
  })
}

export function rangeReadFormulaGridScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  inputCols: number,
  formulaCols: number,
): WorkPaperUniverScenario {
  return rangeReadScenario({
    executeUniverRead: (runtime) => runtime.sheet.getRange(0, inputCols, rowCount, formulaCols).getValues(),
    executeWorkPaperRead: (workbook, sheetId) => workbook.getRangeValues(range(sheetId, 0, inputCols, rowCount - 1, inputCols + formulaCols - 1)),
    family: 'range-read',
    prepareUniverRead: (runtime) =>
      waitForNonNullCells(runtime, [
        { col: inputCols, row: 0 },
        { col: inputCols + formulaCols - 1, row: rowCount - 1 },
      ]),
    rowCount,
    sheet: buildFormulaGridSheet(rowCount, inputCols, formulaCols),
    summarizeUniverRead: (values) => summarizeFormulaGridRead(values, normalizeBenchmarkValue),
    summarizeWorkPaperRead: (values) => summarizeFormulaGridRead(values, normalizeWorkPaperValue),
    workload,
  })
}

export function indexedLookupAfterColumnWriteScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 0, row: rowCount, value: rowCount + 1_000 },
    family: 'lookup-after-write',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildLookupSheet(rowCount),
    workbookOptions: { useColumnIndex: true },
    workload,
  })
}

export function indexedLookupAfterBatchWriteScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  editCount: number,
): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => {
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        runtime.sheet.getRange(row, 0).setValue(row + 10_000)
      }
    },
    executeWorkPaperMutation: (workbook) => {
      const sheetId = workbook.getSheetId('Bench')!
      workbook.suspendEvaluation()
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        workbook.setCellContents(address(sheetId, row, 0), row + 10_000)
      }
      return workbook.resumeEvaluation()
    },
    family: 'lookup-after-write',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildLookupSheet(rowCount),
    workbookOptions: { useColumnIndex: true },
    workload,
  })
}

export function lookupApproximateAfterColumnWriteScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 0, row: rowCount, value: rowCount + 1 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupSheet(rowCount),
    workload,
  })
}

export function lookupApproximateDescendingScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 3, row: 0, value: Math.floor(rowCount / 3) + 0.5 },
    family: 'lookup-approximate',
    observedCells: [{ col: 4, key: 'formulaValue', row: 0 }],
    rowCount: rowCount + 1,
    sheet: buildApproxLookupDescendingSheet(rowCount),
    workload,
  })
}

export function conditionalAggregationScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  formulaCopies: number,
): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 1, row: rowCount, value: rowCount * 2 },
    family: 'conditional-aggregation',
    observedCells: [
      { col: 4, key: 'sumifValue', row: 0 },
      { col: 4 + formulaCopies, key: 'countifValue', row: 0 },
    ],
    rowCount: rowCount + 1,
    sheet: buildConditionalAggregationSheet(rowCount, formulaCopies),
    workload,
  })
}

export function conditionalAggregationCriteriaEditScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  formulaCopies: number,
): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 3, row: 0, value: 'B' },
    family: 'conditional-aggregation',
    observedCells: [
      { col: 4, key: 'sumifValue', row: 0 },
      { col: 4 + formulaCopies, key: 'countifValue', row: 0 },
    ],
    rowCount: rowCount + 1,
    sheet: buildConditionalAggregationSheet(rowCount, formulaCopies),
    workload,
  })
}

export function conditionalAggregationSharedCriteriaScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  criteriaCount: number,
): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 3, row: 0, value: 'B' },
    family: 'conditional-aggregation',
    observedCells: [
      { col: 3 + criteriaCount, key: 'firstSum', row: 0 },
      { col: 3 + criteriaCount * 2 - 1, key: 'lastSum', row: 0 },
    ],
    rowCount: rowCount + 1,
    sheet: buildConditionalAggregationSharedCriteriaSheet(rowCount, criteriaCount),
    workload,
  })
}

export function conditionalAggregationMixedCriteriaScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  formulaCopies: number,
): WorkPaperUniverScenario {
  return operationScenario({
    edit: { col: 4, row: 0, value: 20 },
    family: 'conditional-aggregation',
    observedCells: [
      { col: 5, key: 'firstCount', row: 0 },
      { col: 5 + formulaCopies, key: 'firstSum', row: 0 },
    ],
    rowCount: rowCount + 1,
    sheet: buildConditionalAggregationMixedSheet(rowCount, formulaCopies),
    workload,
  })
}

export function operationScenario(args: {
  readonly edit?: { readonly col: number; readonly row: number; readonly value: UniverEditableValue }
  readonly executeUniverMutation?: WorkPaperUniverScenario['executeUniverMutation']
  readonly executeWorkPaperMutation?: WorkPaperUniverScenario['executeWorkPaperMutation']
  readonly family: WorkPaperUniverWorkloadFamily
  readonly observedCells: readonly (ObservedCell & { readonly value?: number })[]
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly workbookOptions?: Parameters<typeof WorkPaper.buildFromSheets>[1]
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const sheetName = 'Bench'
  const columnCount = Math.max(1, ...(args.sheet.map((row) => row.length)), ...args.observedCells.map((cell) => cell.col + 1))
  const fixture = {
    ...(args.edit
      ? {
          edit: {
            address: formatA1(args.edit.row, args.edit.col),
            col: args.edit.col,
            row: args.edit.row,
            sheetName,
            value: args.edit.value,
          },
        }
      : {}),
    family: args.family,
    formula: firstFormula(args.sheet) ?? args.workload,
    result: {
      address: formatA1(args.observedCells[0]!.row, args.observedCells[0]!.col),
      col: args.observedCells[0]!.col,
      row: args.observedCells[0]!.row,
      sheetName,
    },
    columnCount,
    rowCount: args.rowCount,
  }
  return {
    kind: 'mutation',
    fixture,
    buildWorkPaperSheets: () => ({ [sheetName]: args.sheet }),
    ...(args.executeUniverMutation ? { executeUniverMutation: args.executeUniverMutation } : {}),
    ...(args.executeWorkPaperMutation ? { executeWorkPaperMutation: args.executeWorkPaperMutation } : {}),
    ...(args.workbookOptions ? { workpaperOptions: args.workbookOptions } : {}),
    setupUniver: (runtime) => setupUniverSheet(runtime, args.sheet),
    verifyUniver: (runtime) => observeUniver(runtime, args.observedCells),
    verifyWorkPaper: (workbook) => observeWorkPaper(workbook, args.observedCells),
  }
}

function rangeReadScenario(args: {
  readonly executeUniverRead: (runtime: UniverRuntime) => unknown[][]
  readonly executeWorkPaperRead: (workbook: WorkPaper, sheetId: number) => unknown[][]
  readonly family: WorkPaperUniverWorkloadFamily
  readonly prepareUniverRead?: (runtime: UniverRuntime) => Promise<void>
  readonly rowCount: number
  readonly sheet: WorkPaperSheet
  readonly summarizeUniverRead: (values: unknown[][]) => Record<string, unknown>
  readonly summarizeWorkPaperRead: (values: unknown[][]) => Record<string, unknown>
  readonly workload: WorkPaperUniverWorkload
}): WorkPaperUniverScenario {
  const scenario = operationScenario({
    executeUniverMutation: (runtime) => args.executeUniverRead(runtime),
    executeWorkPaperMutation: (workbook) => args.executeWorkPaperRead(workbook, workbook.getSheetId('Bench')!),
    family: args.family,
    observedCells: [{ col: 0, key: 'range', row: 0 }],
    rowCount: args.rowCount,
    sheet: args.sheet,
    workload: args.workload,
  })
  return {
    ...scenario,
    ...(args.prepareUniverRead ? { prepareUniverOperation: args.prepareUniverRead } : {}),
    verifyUniver: (_runtime, operationResult) => args.summarizeUniverRead(toGrid(operationResult)),
    verifyWorkPaper: (_workbook, operationResult) => args.summarizeWorkPaperRead(toGrid(operationResult)),
  }
}

async function waitForNonNullCells(
  runtime: UniverRuntime,
  cells: readonly { readonly col: number; readonly row: number }[],
): Promise<void> {
  const deadline = performance.now() + 10_000
  while (cells.some((cell) => runtime.sheet.getRange(cell.row, cell.col).getValue() === null)) {
    if (performance.now() >= deadline) {
      throw new Error('Timed out waiting for Univer formulas before range-read workload')
    }
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Setup polling waits for formula cells before timed range read starts.
    await new Promise((resolve) => setTimeout(resolve, 1))
  }
}

async function setupUniverSheet(runtime: UniverRuntime, sheet: WorkPaperSheet): Promise<void> {
  const rowCount = sheet.length
  const columnCount = Math.max(1, ...sheet.map((row) => row.length))
  const values = sheet.map((row) => Array.from({ length: columnCount }, (_value, col) => nonFormulaValue(row[col])))
  runtime.sheet.getRange(0, 0, rowCount, columnCount).setValues(values)
  for (const run of collectFormulaRuns(sheet)) {
    runtime.sheet.getRange(run.row, run.startCol, 1, run.formulas.length).setFormulas([run.formulas])
  }
}

function rectangularEditValues(rowCount: number, inputCols: number): number[][] {
  return Array.from({ length: rowCount }, (_rowValue, row) =>
    Array.from({ length: inputCols }, (_colValue, col) => (row + 1) * (col + 2)),
  )
}

function emptyGrid(rowCount: number, colCount: number): string[][] {
  return Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => ''))
}

function collectFormulaRuns(sheet: WorkPaperSheet): { readonly formulas: string[]; readonly row: number; readonly startCol: number }[] {
  const formulaRuns: { readonly formulas: string[]; readonly row: number; readonly startCol: number }[] = []
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
  return formulaRuns
}

function nonFormulaValue(value: WorkPaperSheet[number][number] | undefined): UniverEditableValue {
  return typeof value === 'string' && value.startsWith('=') ? '' : value ?? ''
}

function observeWorkPaper(workbook: WorkPaper, observedCells: readonly (ObservedCell & { readonly value?: number })[]): Record<string, unknown> {
  const sheetId = workbook.getSheetId('Bench')!
  return Object.fromEntries(
    observedCells.map((cell) => [
      cell.key,
      cell.value ?? normalizeBenchmarkValue(normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, cell.row, cell.col)))),
    ]),
  )
}

function observeUniver(runtime: UniverRuntime, observedCells: readonly (ObservedCell & { readonly value?: number })[]): Record<string, unknown> {
  return Object.fromEntries(
    observedCells.map((cell) => [
      cell.key,
      cell.value ?? normalizeBenchmarkValue(runtime.sheet.getRange(formatA1(cell.row, cell.col)).getValue()),
    ]),
  )
}

function summarizeDenseRangeRead(values: unknown[][], normalize: (value: unknown) => unknown): Record<string, unknown> {
  const lastRow = values.at(-1)
  return {
    readCols: values[0]?.length ?? 0,
    readRows: values.length,
    terminalValue: normalize(lastRow?.at(-1)),
    topLeftValue: normalize(values[0]?.[0]),
  }
}

function summarizeSparseWideRead(
  values: unknown[][],
  middleCol: number,
  normalize: (value: unknown) => unknown,
): Record<string, unknown> {
  const lastRow = values.at(-1)
  return {
    emptyValue: normalize(values[0]?.[1]) ?? '',
    middleValue: normalize(lastRow?.[middleCol]),
    readCols: values[0]?.length ?? 0,
    readRows: values.length,
    terminalValue: normalize(lastRow?.at(-1)),
    topLeftValue: normalize(values[0]?.[0]),
  }
}

function summarizeFormulaGridRead(values: unknown[][], normalize: (value: unknown) => unknown): Record<string, unknown> {
  const lastRow = values.at(-1)
  return {
    leadingFormulaValue: normalize(values[0]?.[0]),
    readCols: values[0]?.length ?? 0,
    readRows: values.length,
    terminalFormulaValue: normalize(lastRow?.at(-1)),
  }
}

function toGrid(value: unknown): unknown[][] {
  if (Array.isArray(value) && value.every(Array.isArray)) {
    return value
  }
  throw new Error('Expected Univer range-read operation to return a rectangular value grid')
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
