import { WorkPaper } from '@bilig/headless'
import {
  address,
  buildApproxLookupDescendingSheet,
  buildApproxLookupDuplicateSheet,
  buildApproxLookupSheet,
  buildLookupSheet,
} from './workpaper-benchmark-fixtures.js'
import {
  HyperFormula,
  measureHyperFormulaMutationSample,
  measureMutationSample,
  normalizeHyperFormulaValue,
  normalizeWorkPaperValue,
  toHyperFormulaSheet,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { HYPERFORMULA_LICENSE_KEY } from './benchmark-workpaper-vs-hyperformula.js'

export function measureWorkPaperIndexedLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex: true })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexedLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: true },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperIndexedLookupAfterBatchWriteSample(rowCount: number, editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex: true })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        workbook.setCellContents(address(sheetId, row, 0), row + 10_000)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexedLookupAfterBatchWriteSample(rowCount: number, editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: true },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        workbook.setCellContents(address(sheetId, row, 0), row + 10_000)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperApproximateLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaApproximateLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperApproximateLookupDescendingSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupDescendingSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 3) + 0.5),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaApproximateLookupDescendingSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupDescendingSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 3) + 0.5),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperApproximateLookupDuplicateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupDuplicateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 5)),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaApproximateLookupDuplicateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupDuplicateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 5)),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}
