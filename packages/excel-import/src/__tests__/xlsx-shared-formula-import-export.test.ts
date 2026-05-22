import { strFromU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { buildSharedFormulaWorkbookBytes } from '@bilig/excel-fixtures'
import { exportXlsx, importXlsx } from '../index.js'

describe('XLSX shared formula import/export', () => {
  it('expands Excel shared formulas into canonical cell formulas and exports them without stale shared anchors', () => {
    const imported = importXlsx(buildSharedFormulaWorkbookBytes(), 'shared-formula-model.xlsx')
    const sheet = imported.snapshot.sheets.find((candidate) => candidate.name === 'Model')
    const cells = new Map(sheet?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cells.get('B2')).toMatchObject({ formula: 'A2*2', value: 20 })
    expect(cells.get('B3')).toMatchObject({ formula: 'A3*2', value: 40 })
    expect(cells.get('B4')).toMatchObject({ formula: 'A4*2', value: 60 })
    expect(cells.get('C2')).toMatchObject({ formula: 'B2+1', value: 21 })
    expect(cells.get('C3')).toMatchObject({ formula: 'B3+1', value: 41 })
    expect(cells.get('C4')).toMatchObject({ formula: 'B4+1', value: 61 })

    const exportedSheetXml = strFromU8(unzipSync(exportXlsx(imported.snapshot))['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(exportedSheetXml).not.toContain('t="shared"')
    expect(cellXml(exportedSheetXml, 'B3')).toContain('<f>A3*2</f>')
    expect(cellXml(exportedSheetXml, 'C4')).toContain('<f>B4+1</f>')

    const roundTrip = importXlsx(exportXlsx(imported.snapshot), 'shared-formula-model-roundtrip.xlsx')
    const roundTripCells = new Map(roundTrip.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    expect(roundTripCells.get('B3')).toMatchObject({ formula: 'A3*2', value: 40 })
    expect(roundTripCells.get('C4')).toMatchObject({ formula: 'B4+1', value: 61 })
  })
})

function cellXml(sheetXml: string, address: string): string {
  return sheetXml.match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
}
