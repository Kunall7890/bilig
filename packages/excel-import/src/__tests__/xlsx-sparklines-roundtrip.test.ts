import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const x14Namespace = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const xmNamespace = 'http://schemas.microsoft.com/office/excel/2006/main'
const sparklineExtensionXml = `<ext uri="${sparklineExtensionUri}" xmlns:x14="${x14Namespace}"><x14:sparklineGroups xmlns:xm="${xmNamespace}"><x14:sparklineGroup type="line" displayEmptyCellsAs="gap" markers="1"><x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FF376092"/><x14:colorFirst rgb="FF376092"/><x14:colorLast rgb="FF376092"/><x14:colorHigh rgb="FF376092"/><x14:colorLow rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>Data!A2:D2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>Data!A3:D3</xm:f><xm:sqref>E3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`
const rootNamespacedSparklineExtensionXml = `<ext uri="${sparklineExtensionUri}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Data!A2:D2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`

describe('worksheet sparkline metadata import/export', () => {
  it('preserves worksheet sparkline extension metadata across XLSX round trips', () => {
    const sourceBytes = buildWorkbookWithSparklineExtension()
    const imported = importXlsx(sourceBytes, 'sparklines.xlsx')

    const exportedSheetXml = worksheetXml(exportXlsx(imported.snapshot), 1)

    expect(imported.warnings).toEqual([])
    expect(sparklineGroupCount(sourceBytes)).toBe(1)
    expect(sparklineCount(sourceBytes)).toBe(2)
    expect(exportedSheetXml).toContain(`<ext uri="${sparklineExtensionUri}"`)
    expect(sparklineGroupCount(exportXlsx(imported.snapshot))).toBe(1)
    expect(sparklineCount(exportXlsx(imported.snapshot))).toBe(2)
  })

  it('keeps worksheet namespace declarations needed by preserved sparkline XML', () => {
    const sourceBytes = buildWorkbookWithRootNamespacedSparklineExtension()
    const imported = importXlsx(sourceBytes, 'sparklines-root-namespace.xlsx')

    const exportedSheetXml = worksheetXml(exportXlsx(imported.snapshot), 1)

    expect(exportedSheetXml).toContain(`xmlns:x14="${x14Namespace}"`)
    expect(exportedSheetXml).toContain(`xmlns:xm="${xmNamespace}"`)
    expect(exportedSheetXml).toContain('<x14:sparklineGroups>')
    expect(exportedSheetXml).toContain('<xm:f>Data!A2:D2</xm:f>')
  })

  it('does not add worksheet sparkline extensions to workbooks without sparklines', () => {
    const exportedSheetXml = worksheetXml(exportXlsx(buildSparklineWorkbook()), 1)

    expect(exportedSheetXml).not.toContain('sparklineGroups')
    expect(exportedSheetXml).not.toContain(sparklineExtensionUri)
  })
})

function buildWorkbookWithSparklineExtension(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildSparklineWorkbook()))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(sheetXml.replace('</worksheet>', `<extLst>${sparklineExtensionXml}</extLst></worksheet>`))
  return zipSync(zip)
}

function buildWorkbookWithRootNamespacedSparklineExtension(): Uint8Array {
  const zip = unzipSync(exportXlsx(buildSparklineWorkbook()))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(
    sheetXml
      .replace('<worksheet ', `<worksheet xmlns:x14="${x14Namespace}" xmlns:xm="${xmNamespace}" `)
      .replace('</worksheet>', `<extLst>${rootNamespacedSparklineExtensionXml}</extLst></worksheet>`),
  )
  return zipSync(zip)
}

function buildSparklineWorkbook(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'sparklines',
    },
    sheets: [
      {
        id: 1,
        name: 'Data',
        order: 0,
        cells: [
          { address: 'A1', value: 'Q1' },
          { address: 'B1', value: 'Q2' },
          { address: 'C1', value: 'Q3' },
          { address: 'D1', value: 'Q4' },
          { address: 'E1', value: 'Trend' },
          { address: 'A2', value: 10 },
          { address: 'B2', value: 20 },
          { address: 'C2', value: 15 },
          { address: 'D2', value: 30 },
          { address: 'E2', value: '' },
          { address: 'A3', value: 18 },
          { address: 'B3', value: 12 },
          { address: 'C3', value: 24 },
          { address: 'D3', value: 28 },
          { address: 'E3', value: '' },
        ],
      },
    ],
  }
}

function sparklineGroupCount(bytes: Uint8Array): number {
  return worksheetXml(bytes, 1).match(/<x14:sparklineGroup\b/gu)?.length ?? 0
}

function sparklineCount(bytes: Uint8Array): number {
  return worksheetXml(bytes, 1).match(/<x14:sparkline\b/gu)?.length ?? 0
}

function worksheetXml(bytes: Uint8Array, sheetIndex: number): string {
  return strFromU8(unzipSync(bytes)[`xl/worksheets/sheet${String(sheetIndex)}.xml`] ?? new Uint8Array())
}
