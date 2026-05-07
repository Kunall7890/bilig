import { WorkPaper } from '@bilig/headless'
import { HYPERFORMULA_LICENSE_KEY } from './benchmark-workpaper-vs-hyperformula.js'
import {
  address,
  buildDynamicArraySortSheet,
  buildDynamicArrayUniqueSheet,
  buildLookupSearchModeReverseSheet,
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

export function measureWorkPaperReverseSearchLookupSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSearchModeReverseSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 2)),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaReverseSearchLookupSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSearchModeReverseSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), Math.floor(rowCount / 2)),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperDynamicArraySortSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDynamicArraySortSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const spillAnchor = address(sheetId, 0, 1)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), 0),
    () => ({
      spillHeight: workbook.getSheetDimensions(sheetId).height,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillValue: normalizeWorkPaperValue(workbook.getCellValue(spillAnchor)),
    }),
  )
}

export function measureWorkPaperDynamicArrayUniqueSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDynamicArrayUniqueSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const spillAnchor = address(sheetId, 0, 1)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount),
    () => ({
      spillHeight: workbook.getSheetDimensions(sheetId).height,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillValue: normalizeWorkPaperValue(workbook.getCellValue(spillAnchor)),
    }),
  )
}
