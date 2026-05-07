import { WorkPaper } from '../../headless/src/work-paper.js'
import { HYPERFORMULA_LICENSE_KEY } from './benchmark-workpaper-vs-hyperformula.js'
import {
  address,
  buildApproxLookupSheet,
  buildBatchMultiColumnRows,
  buildDenseLiteralSheet,
  buildDynamicArraySheet,
  buildFormulaChainRow,
  buildFormulaEditChainRow,
  buildFormulaFanoutRow,
  buildLookupSheet,
  buildMixedContentSheet,
  buildMultiSheetLiteralSheets,
  buildTextLookupSheet,
  buildValueFormulaRows,
  range,
} from './workpaper-benchmark-fixtures.js'
import {
  HyperFormula,
  measureHyperFormulaBuildFromSheets,
  measureHyperFormulaMutationSample,
  measureMutationSample,
  measureWorkPaperBuildFromSheets,
  normalizeHyperFormulaValue,
  normalizeWorkPaperValue,
  toHyperFormulaSheet,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'

export function measureWorkPaperDenseBuildSample(rows: number, cols: number): BenchmarkSample {
  const sheet = buildDenseLiteralSheet(rows, cols)
  return measureWorkPaperBuildFromSheets({ Bench: sheet }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    }
  })
}

export function measureHyperFormulaDenseBuildSample(rows: number, cols: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets({ Bench: toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols)) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    }
  })
}

export function measureWorkPaperMixedBuildSample(rowCount: number): BenchmarkSample {
  const sheet = buildMixedContentSheet(rowCount)
  return measureWorkPaperBuildFromSheets({ Bench: sheet }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 5))),
    }
  })
}

export function measureHyperFormulaMixedBuildSample(rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets({ Bench: toHyperFormulaSheet(buildMixedContentSheet(rowCount)) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 5))),
    }
  })
}

export function measureWorkPaperManySheetsBuildSample(sheetCount: number, rows: number, cols: number): BenchmarkSample {
  const sheets = buildMultiSheetLiteralSheets(sheetCount, rows, cols)
  return measureWorkPaperBuildFromSheets(sheets, (workbook) => {
    const sheetId = workbook.getSheetId(`Sheet${sheetCount}`)!
    return {
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    }
  })
}

export function measureHyperFormulaManySheetsBuildSample(sheetCount: number, rows: number, cols: number): BenchmarkSample {
  const sheets = Object.fromEntries(
    Object.entries(buildMultiSheetLiteralSheets(sheetCount, rows, cols)).map(([sheetName, sheet]) => [
      sheetName,
      toHyperFormulaSheet(sheet),
    ]),
  )
  return measureHyperFormulaBuildFromSheets(sheets, (workbook) => {
    const sheetId = workbook.getSheetId(`Sheet${sheetCount}`)!
    return {
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    }
  })
}

export function measureWorkPaperLegacySingleEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: [buildFormulaChainRow(downstreamCount)] })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? null,
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
    }),
  )
}

export function measureHyperFormulaLegacySingleEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaChainRow(downstreamCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? null,
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
    }),
  )
}

export function measureWorkPaperSingleChainEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: [buildFormulaChainRow(downstreamCount)] })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
    }),
  )
}

export function measureHyperFormulaSingleChainEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaChainRow(downstreamCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, downstreamCount))),
    }),
  )
}

export function measureWorkPaperSingleFanoutEditSample(fanoutCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: [buildFormulaFanoutRow(fanoutCount)] })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, fanoutCount))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureHyperFormulaSingleFanoutEditSample(fanoutCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaFanoutRow(fanoutCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, fanoutCount))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureWorkPaperFormulaEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: [buildFormulaEditChainRow(downstreamCount)],
  })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 2), '=A1*B1'),
    () => ({
      editedFormula: workbook.getCellFormula(address(sheetId, 0, 2)) ?? null,
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, downstreamCount + 2))),
    }),
  )
}

export function measureHyperFormulaFormulaEditSample(downstreamCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet([buildFormulaEditChainRow(downstreamCount)]) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 2), '=A1*B1'),
    () => ({
      editedFormula: workbook.getCellFormula(address(sheetId, 0, 2)) ?? null,
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, downstreamCount + 2))),
    }),
  )
}

export function measureWorkPaperLegacyBatchEditSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      }),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      sampleFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    }),
  )
}

export function measureHyperFormulaLegacyBatchEditSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      }),
    (changes) => ({
      changeCount: Array.isArray(changes) ? changes.length : 0,
      sampleFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    }),
  )
}

export function measureWorkPaperBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      }),
    () => ({
      sampleFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureHyperFormulaBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      }),
    () => ({
      sampleFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureWorkPaperBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildBatchMultiColumnRows(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
          workbook.setCellContents(address(sheetId, row, 1), row * 5)
        }
      }),
    () => ({
      sampleSumValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      sampleProductValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureHyperFormulaBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildBatchMultiColumnRows(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
          workbook.setCellContents(address(sheetId, row, 1), row * 5)
        }
      }),
    () => ({
      sampleSumValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      sampleProductValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureWorkPaperRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseLiteralSheet(rows, cols) })
  const sheetId = workbook.getSheetId('Bench')!
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1)
  return measureMutationSample(
    workbook,
    () => workbook.getRangeValues(targetRange),
    (values) => {
      const lastRow = values.at(-1)
      return {
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeWorkPaperValue(lastRow?.at(-1)),
        topLeftValue: normalizeWorkPaperValue(values[0]?.[0]),
      }
    },
  )
}

export function measureHyperFormulaRangeReadSample(rows: number, cols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildDenseLiteralSheet(rows, cols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetRange = range(sheetId, 0, 0, rows - 1, cols - 1)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.getRangeValues(targetRange),
    (values) => {
      const lastRow = values.at(-1)
      return {
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeHyperFormulaValue(lastRow?.at(-1)),
        topLeftValue: normalizeHyperFormulaValue(values[0]?.[0]),
      }
    },
  )
}

export function measureWorkPaperLookupSample(rowCount: number, useColumnIndex: boolean): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex })
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureHyperFormulaLookupSample(rowCount: number, useColumnIndex: boolean): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureWorkPaperApproximateLookupSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount - 0.5),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureHyperFormulaApproximateLookupSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, rowCount - 0.5),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureWorkPaperTextLookupSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildTextLookupSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, 'KEY-04999'),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureHyperFormulaTextLookupSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildTextLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(targetAddress, 'KEY-04999'),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(formulaAddress)),
    }),
  )
}

export function measureWorkPaperDynamicArraySample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDynamicArraySheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const thresholdAddress = address(sheetId, 0, 1)
  const spillAnchor = address(sheetId, 0, 2)
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(thresholdAddress, rowCount - 10),
    () => ({
      spillHeight: workbook.getSheetDimensions(sheetId).height,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillValue: normalizeWorkPaperValue(workbook.getCellValue(spillAnchor)),
    }),
  )
}
