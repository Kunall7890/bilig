import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

const sparklineExtensionUri = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
const sparklineXml = `<ext uri="${sparklineExtensionUri}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup type="line"><x14:sparklines><x14:sparkline><xm:f>Data!A2:D2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>Data!A3:D3</xm:f><xm:sqref>E3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`

describe('SpreadsheetEngine sparkline metadata', () => {
  it('roundtrips worksheet sparkline extension XML through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sparklines-runtime' })
    await engine.ready()

    engine.importSnapshot(sparklineSnapshot())
    engine.setCellValue('Data', 'F1', 'agent edit')

    const exported = engine.exportSnapshot()
    expect(exported.sheets[0]?.metadata?.sparklines).toEqual({ xml: sparklineXml })

    const restored = new SpreadsheetEngine({ workbookName: 'sparklines-runtime-restored' })
    await restored.ready()
    restored.importSnapshot(exported)

    expect(restored.exportSnapshot().sheets[0]?.metadata?.sparklines).toEqual({ xml: sparklineXml })
  })

  it('rewrites sparkline source and output refs through structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sparklines-structural' })
    await engine.ready()
    engine.importSnapshot(sparklineSnapshot())

    engine.insertRows('Data', 1, 1)

    expect(sparklineRefs(engine.exportSnapshot().sheets[0]?.metadata?.sparklines?.xml)).toEqual([
      { formula: 'Data!A3:D3', sqref: 'E3' },
      { formula: 'Data!A4:D4', sqref: 'E4' },
    ])
  })
})

function sparklineRefs(xml: string | undefined): Array<{ readonly formula: string; readonly sqref: string }> {
  if (!xml) {
    throw new Error('Expected sparkline XML')
  }
  const matches = [...xml.matchAll(/<x14:sparkline><xm:f>([\s\S]*?)<\/xm:f><xm:sqref>([\s\S]*?)<\/xm:sqref><\/x14:sparkline>/gu)]
  return matches.map((match) => ({
    formula: match[1] ?? '',
    sqref: match[2] ?? '',
  }))
}

function sparklineSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Sparklines' },
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
        metadata: {
          sparklines: { xml: sparklineXml },
        },
      },
    ],
  }
}
