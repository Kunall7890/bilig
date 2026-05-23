import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const sheetPr = {
  xml: '<sheetPr codeName="Sheet8"><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
}

describe('engine worksheet sheetPr properties', () => {
  it('roundtrips raw worksheet sheetPr properties through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sheet-properties-runtime' })
    await engine.ready()

    engine.importSnapshot(sheetPropertiesSnapshot())

    expect(engine.exportSnapshot().sheets[0]?.metadata?.sheetPr).toEqual(sheetPr)

    const restored = new SpreadsheetEngine({ workbookName: 'sheet-properties-restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.exportSnapshot().sheets[0]?.metadata?.sheetPr).toEqual(sheetPr)
  })

  it('preserves raw worksheet sheetPr properties across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sheet-properties-structure' })
    await engine.ready()
    engine.importSnapshot(sheetPropertiesSnapshot())

    engine.insertRows('Report', 0, 1)
    engine.insertColumns('Report', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.sheetPr).toEqual(sheetPr)
  })
})

function sheetPropertiesSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Worksheet properties runtime' },
    sheets: [
      {
        id: 1,
        name: 'Report',
        order: 0,
        metadata: {
          sheetPr,
          tabColor: { rgb: 'FFFF0000' },
        },
        cells: [
          { address: 'A1', value: 'Metric' },
          { address: 'A2', value: 'Revenue' },
        ],
      },
    ],
  }
}
