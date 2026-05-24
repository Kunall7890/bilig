import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx, precisionAsDisplayedCalculationWarning } from '../index.js'

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
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet(
    options.formula
      ? [
          ['rate', 'gross'],
          [0.08, { f: '1+A2', v: 1.08 }],
        ]
      : [['rate'], [0.08]],
  )
  XLSX.utils.book_append_sheet(workbook, sheet, 'Model')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sourceWorkbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  const workbookXmlWithoutCalcPr = sourceWorkbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, '')
  zip['xl/workbook.xml'] = strToU8(
    workbookXmlWithoutCalcPr.replace(
      '</workbook>',
      '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="10200" iterateDelta="9.9999999999999995E-7" fullCalcOnLoad="1" calcOnSave="1" calcCompleted="0" concurrentCalc="0"/></workbook>',
    ),
  )
  return zipSync(zip)
}

function workbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}
