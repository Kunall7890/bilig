import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx, precisionAsDisplayedCalculationWarning } from '../index.js'
import { patchXlsxTestZipText, readXlsxTestZipText } from './xlsx-test-helpers.js'

describe('XLSX calculation properties roundtrip', () => {
  it('preserves semantic workbook calcPr attributes through import, export, and engine snapshots', () => {
    const imported = importXlsx(buildCalculationPropertiesWorkbookBytes(), 'calculation-properties.xlsx')

    expect(imported.snapshot.workbook.metadata?.calculationSettings).toEqual({
      mode: 'manual',
      compatibilityMode: 'excel-modern',
      fullPrecision: false,
      iterate: true,
      iterateCount: 10200,
      iterateDelta: '9.9999999999999995E-7',
      fullCalcOnLoad: true,
      calcOnSave: true,
      calcCompleted: false,
      concurrentCalc: false,
    })
    expect(imported.warnings).toEqual([expect.stringContaining('Manual calculation mode')])

    const exportedWorkbookXml = workbookXml(exportXlsx(imported.snapshot))
    expect(exportedWorkbookXml).toContain(
      '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" calcOnSave="1" calcCompleted="0" concurrentCalc="0"/>',
    )

    const engine = new SpreadsheetEngine({ workbookName: 'calculation-properties-engine' })
    engine.importSnapshot(imported.snapshot)
    const exportedFromEngineWorkbookXml = workbookXml(exportXlsx(engine.exportSnapshot()))
    expect(exportedFromEngineWorkbookXml).toContain(
      '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" calcOnSave="1" calcCompleted="0" concurrentCalc="0"/>',
    )
  })

  it('warns for precision-as-displayed workbooks only when formulas need recalculation semantics', () => {
    const staticImported = importXlsx(buildCalculationPropertiesWorkbookBytes({ formula: false }), 'static-precision.xlsx')
    expect(staticImported.snapshot.workbook.metadata?.calculationSettings).toMatchObject({ fullPrecision: false })
    expect(staticImported.warnings).not.toContain(precisionAsDisplayedCalculationWarning)

    const formulaImported = importXlsx(buildCalculationPropertiesWorkbookBytes({ formula: true }), 'formula-precision.xlsx')
    expect(formulaImported.snapshot.workbook.metadata?.calculationSettings).toMatchObject({ fullPrecision: false })
    expect(formulaImported.warnings).toContain(precisionAsDisplayedCalculationWarning)
  })
})

function buildCalculationPropertiesWorkbookBytes(options: { readonly formula?: boolean } = {}): Uint8Array {
  const cells = options.formula
    ? [
        { address: 'A1', row: 0, col: 0, value: 'rate' },
        { address: 'B1', row: 0, col: 1, value: 'gross' },
        { address: 'A2', row: 1, col: 0, value: 0.08 },
        { address: 'B2', row: 1, col: 1, formula: '1+A2', value: 1.08 },
      ]
    : [
        { address: 'A1', row: 0, col: 0, value: 'rate' },
        { address: 'A2', row: 1, col: 0, value: 0.08 },
      ]
  return patchXlsxTestZipText(
    writeSimpleXlsxWorkbook({
      sheets: [{ name: 'Model', cells }],
    }),
    'xl/workbook.xml',
    (sourceWorkbookXml) =>
      sourceWorkbookXml
        .replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, '')
        .replace(
          '</workbook>',
          '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" calcOnSave="1" calcCompleted="0" concurrentCalc="0"/></workbook>',
        ),
  )
}

function workbookXml(bytes: Uint8Array): string {
  return readXlsxTestZipText(bytes, 'xl/workbook.xml')
}
