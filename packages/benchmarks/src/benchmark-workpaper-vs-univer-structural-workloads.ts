import type { WorkPaper } from '../../headless/src/work-paper.js'
import type { WorkPaperUniverScenario, WorkPaperUniverWorkload } from './benchmark-workpaper-vs-univer.js'
import { operationScenario } from './benchmark-workpaper-vs-univer-operation-workloads.js'
import {
  address,
  buildOverlappingAggregateSheet,
  buildRectangularBlockFormulaRows,
  buildRectangularBlockFormulaSheet,
  buildStructuralColumnSheet,
} from './workpaper-benchmark-fixtures.js'

export function structuralInsertRowsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const insertIndex = Math.floor(rowCount / 2)
  return withDimensions(
    operationScenario({
      executeUniverMutation: (runtime) => runtime.sheet.insertRowsBefore(insertIndex, 1),
      executeWorkPaperMutation: (workbook) => workbook.addRows(workbook.getSheetId('Bench')!, insertIndex, 1),
      family: 'structural-rows',
      observedCells: [{ col: 1, key: 'terminalSum', row: rowCount }],
      rowCount,
      sheet: buildOverlappingAggregateSheet(rowCount),
      workload,
    }),
  )
}

export function structuralDeleteRowsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const deleteIndex = Math.floor(rowCount / 2)
  return withDimensions(
    operationScenario({
      executeUniverMutation: (runtime) => runtime.sheet.deleteRows(deleteIndex, 1),
      executeWorkPaperMutation: (workbook) => workbook.removeRows(workbook.getSheetId('Bench')!, deleteIndex, 1),
      family: 'structural-rows',
      observedCells: [{ col: 1, key: 'terminalSum', row: rowCount - 2 }],
      rowCount,
      sheet: buildOverlappingAggregateSheet(rowCount),
      workload,
    }),
  )
}

export function structuralMoveRowsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  const start = Math.floor(rowCount / 2)
  return withDimensions(
    operationScenario({
      executeUniverMutation: (runtime) => runtime.sheet.moveRows(runtime.sheet.getRange(start, 0, 1, runtime.sheet.getMaxColumns()), 0),
      executeWorkPaperMutation: (workbook) => workbook.moveRows(workbook.getSheetId('Bench')!, start, 1, 0),
      family: 'structural-rows',
      observedCells: [{ col: 0, key: 'headValue', row: 0 }],
      rowCount,
      sheet: buildOverlappingAggregateSheet(rowCount),
      workload,
    }),
  )
}

export function structuralInsertColumnsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => runtime.sheet.insertColumnsBefore(1, 1),
    executeWorkPaperMutation: (workbook) => workbook.addColumns(workbook.getSheetId('Bench')!, 1, 1),
    family: 'structural-columns',
    observedCells: [{ col: 4, key: 'terminalFormula', row: rowCount - 1 }],
    rowCount,
    sheet: buildStructuralColumnSheet(rowCount),
    workload,
  })
}

export function structuralDeleteColumnsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => runtime.sheet.deleteColumns(1, 1),
    executeWorkPaperMutation: (workbook) => workbook.removeColumns(workbook.getSheetId('Bench')!, 1, 1),
    family: 'structural-columns',
    observedCells: [{ col: 0, key: 'terminalValue', row: rowCount - 1 }],
    rowCount,
    sheet: buildStructuralColumnSheet(rowCount),
    workload,
  })
}

export function structuralMoveColumnsScenario(workload: WorkPaperUniverWorkload, rowCount: number): WorkPaperUniverScenario {
  return operationScenario({
    executeUniverMutation: (runtime) => runtime.sheet.moveColumns(runtime.sheet.getRange(0, 1, runtime.sheet.getMaxRows(), 1), 0),
    executeWorkPaperMutation: (workbook) => workbook.moveColumns(workbook.getSheetId('Bench')!, 1, 1, 0),
    family: 'structural-columns',
    observedCells: [
      { col: 0, key: 'headValue', row: 0 },
      { col: 3, key: 'terminalFormula', row: rowCount - 1 },
    ],
    rowCount,
    sheet: buildStructuralColumnSheet(rowCount),
    workload,
  })
}

export function appendFormulaRowsScenario(
  workload: WorkPaperUniverWorkload,
  rowCount: number,
  inputCols: number,
  appendCount: number,
): WorkPaperUniverScenario {
  return withDimensions(
    operationScenario({
      executeUniverMutation: (runtime) => {
        runtime.sheet.insertRowsAfter(rowCount - 1, appendCount)
        const appendedRows = buildRectangularBlockFormulaRows(appendCount, inputCols, rowCount + 1)
        runtime.sheet
          .getRange(rowCount, 0, appendCount, inputCols)
          .setValues(appendedRows.map((row) => row.slice(0, inputCols).map((value) => (value === null ? '' : value))))
        runtime.sheet.getRange(rowCount, inputCols, appendCount, 1).setFormulas(appendedRows.map((row) => [String(row[inputCols])]))
      },
      executeWorkPaperMutation: (workbook) => {
        const sheetId = workbook.getSheetId('Bench')!
        return workbook.batch(() => {
          workbook.addRows(sheetId, rowCount, appendCount)
          workbook.setCellContents(address(sheetId, rowCount, 0), buildRectangularBlockFormulaRows(appendCount, inputCols, rowCount + 1))
        })
      },
      family: 'structural-rows',
      observedCells: [
        { col: inputCols, key: 'appendedLeadingSum', row: rowCount },
        { col: inputCols, key: 'appendedTerminalSum', row: rowCount + appendCount - 1 },
      ],
      rowCount,
      sheet: buildRectangularBlockFormulaSheet(rowCount, inputCols),
      workload,
    }),
  )
}

function withDimensions(scenario: WorkPaperUniverScenario): WorkPaperUniverScenario {
  return {
    ...scenario,
    verifyUniver: (runtime, operationResult) => ({
      ...scenario.verifyUniver(runtime, operationResult),
      dimensions: { height: runtime.sheet.getMaxRows(), width: runtime.sheet.getMaxColumns() },
    }),
    verifyWorkPaper: (workbook, operationResult) => ({
      ...scenario.verifyWorkPaper(workbook, operationResult),
      dimensions: dimensionsOf(workbook),
    }),
  }
}

function dimensionsOf(workbook: WorkPaper): { readonly height: number; readonly width: number } {
  const dimensions = workbook.getSheetDimensions(workbook.getSheetId('Bench')!)
  return { height: dimensions.height, width: dimensions.width }
}
