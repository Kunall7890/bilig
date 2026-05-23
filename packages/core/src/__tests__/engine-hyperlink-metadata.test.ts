import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine hyperlink metadata', () => {
  it('roundtrips hyperlinks through snapshot restore and export', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'hyperlink-roundtrip' })
    await engine.ready()

    engine.importSnapshot(hyperlinkSnapshot())

    expect(engine.getHyperlink('Links', 'A1')).toEqual({
      sheetName: 'Links',
      address: 'A1',
      target: 'https://example.com/report',
      tooltip: 'Open report',
      display: 'Report',
    })
    expect(engine.exportSnapshot().sheets[0]?.metadata?.hyperlinks).toEqual([
      {
        sheetName: 'Links',
        address: 'A1',
        target: 'https://example.com/report',
        tooltip: 'Open report',
        display: 'Report',
      },
      {
        sheetName: 'Links',
        address: 'B2',
        target: '#Summary!A1',
        tooltip: 'Jump',
        display: 'Summary',
      },
    ])
  })

  it('retargets and drops hyperlinks through structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'hyperlink-structural' })
    await engine.ready()
    engine.importSnapshot(hyperlinkSnapshot())

    engine.insertRows('Links', 0, 1)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.hyperlinks).toEqual([
      {
        sheetName: 'Links',
        address: 'A2',
        target: 'https://example.com/report',
        tooltip: 'Open report',
        display: 'Report',
      },
      {
        sheetName: 'Links',
        address: 'B3',
        target: '#Summary!A1',
        tooltip: 'Jump',
        display: 'Summary',
      },
    ])

    engine.deleteRows('Links', 1, 1)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.hyperlinks).toEqual([
      {
        sheetName: 'Links',
        address: 'B2',
        target: '#Summary!A1',
        tooltip: 'Jump',
        display: 'Summary',
      },
    ])
  })
})

function hyperlinkSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Hyperlinks' },
    sheets: [
      {
        id: 1,
        name: 'Links',
        order: 0,
        cells: [
          { address: 'A1', value: 'Report' },
          { address: 'B2', value: 'Summary' },
        ],
        metadata: {
          hyperlinks: [
            {
              sheetName: 'Links',
              address: 'A1',
              target: 'https://example.com/report',
              tooltip: 'Open report',
              display: 'Report',
            },
            {
              sheetName: 'Links',
              address: 'B2',
              target: '#Summary!A1',
              tooltip: 'Jump',
              display: 'Summary',
            },
          ],
        },
      },
      {
        id: 2,
        name: 'Summary',
        order: 1,
        cells: [{ address: 'A1', value: 'Destination' }],
      },
    ],
  }
}
