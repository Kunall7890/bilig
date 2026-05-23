import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

const printPageSetup = {
  printOptionsXml: '<printOptions horizontalCentered="1" gridLines="1"/>',
  pageMarginsXml: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
  pageSetupXml: '<pageSetup paperSize="9" scale="60" orientation="landscape"/>',
  headerFooterXml: '<headerFooter alignWithMargins="0"><oddFooter>Page &amp;P</oddFooter></headerFooter>',
  rowBreaksXml: '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="6" max="16383" man="1"/></rowBreaks>',
  colBreaksXml:
    '<colBreaks count="2" manualBreakCount="2"><brk id="2" max="1048575" man="1"/><brk id="5" max="1048575" man="1"/></colBreaks>',
} as const

const printerSettings = [
  {
    relationshipTarget: '../printerSettings/printerSettings1.bin',
    storage: 'base64' as const,
    dataBase64: 'AAECAw==',
    byteLength: 4,
    pageSetupXml: '<pageSetup paperSize="9" r:id="rIdPrinterSettings1"/>',
  },
] as const

describe('SpreadsheetEngine print page setup metadata', () => {
  it('roundtrips print page setup and printer settings through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'print-page-setup-runtime' })
    await engine.ready()

    engine.importSnapshot(printPageSetupSnapshot())
    engine.setCellValue('Report', 'A9', 'agent edit')

    const exported = engine.exportSnapshot()
    expect(exported.sheets[0]?.metadata?.printPageSetup).toEqual(printPageSetup)
    expect(exported.sheets[0]?.metadata?.printerSettings).toEqual(printerSettings)

    const restored = new SpreadsheetEngine({ workbookName: 'print-page-setup-runtime-restored' })
    await restored.ready()
    restored.importSnapshot(exported)

    expect(restored.exportSnapshot().sheets[0]?.metadata?.printPageSetup).toEqual(printPageSetup)
    expect(restored.exportSnapshot().sheets[0]?.metadata?.printerSettings).toEqual(printerSettings)
  })

  it('rewrites manual row breaks through structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'print-page-setup-row-structural' })
    await engine.ready()
    engine.importSnapshot(printPageSetupSnapshot())

    engine.insertRows('Report', 0, 1)

    const setup = engine.exportSnapshot().sheets[0]?.metadata?.printPageSetup
    expect(breakIds(setup?.rowBreaksXml)).toEqual([4, 7])
    expect(breakIds(setup?.colBreaksXml)).toEqual([2, 5])
    expect(setup?.rowBreaksXml).toContain('manualBreakCount="2"')
  })

  it('rewrites manual column breaks through structural column inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'print-page-setup-column-structural' })
    await engine.ready()
    engine.importSnapshot(printPageSetupSnapshot())

    engine.insertColumns('Report', 0, 1)

    const setup = engine.exportSnapshot().sheets[0]?.metadata?.printPageSetup
    expect(breakIds(setup?.rowBreaksXml)).toEqual([3, 6])
    expect(breakIds(setup?.colBreaksXml)).toEqual([3, 6])
    expect(setup?.colBreaksXml).toContain('manualBreakCount="2"')
  })
})

function breakIds(xml: string | undefined): number[] {
  if (!xml) {
    throw new Error('Expected break XML')
  }
  return [...xml.matchAll(/<brk\b[^>]*\bid="(\d+)"/gu)].map((match) => Number(match[1]))
}

function printPageSetupSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Print page setup' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          printPageSetup,
          printerSettings: [...printerSettings],
        },
        cells: Array.from({ length: 8 }, (_value, index) => ({
          address: `A${String(index + 1)}`,
          value: `row-${String(index + 1)}`,
        })),
      },
    ],
  }
}
