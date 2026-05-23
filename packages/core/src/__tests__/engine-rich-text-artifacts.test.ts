import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const richTextArtifacts = {
  cells: [
    {
      address: 'A1',
      text: 'Important: Before signing off',
      storage: 'sharedString' as const,
      xml: '<si><r><rPr><b/></rPr><t>Important:</t></r><r><t xml:space="preserve"> Before signing off</t></r></si>',
    },
    {
      address: 'B2',
      text: 'Revenue sensitivity',
      storage: 'inlineString' as const,
      xml: '<is><r><rPr><u/></rPr><t>Revenue</t></r><r><t xml:space="preserve"> sensitivity</t></r></is>',
    },
  ],
}

describe('engine rich text artifacts', () => {
  it('roundtrips worksheet rich text artifacts through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'rich-text-runtime' })
    await engine.ready()

    engine.importSnapshot(richTextSnapshot())

    expect(engine.exportSnapshot().sheets[0]?.metadata?.richTextArtifacts).toEqual(richTextArtifacts)

    const restored = new SpreadsheetEngine({ workbookName: 'rich-text-runtime-restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.exportSnapshot().sheets[0]?.metadata?.richTextArtifacts).toEqual(richTextArtifacts)
  })

  it('rewrites rich text artifact addresses during structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'rich-text-structure' })
    await engine.ready()
    engine.importSnapshot(richTextSnapshot())

    engine.insertRows('Labels', 0, 1)
    expect(richTextCells(engine.exportSnapshot())).toEqual([
      { address: 'A2', text: 'Important: Before signing off' },
      { address: 'B3', text: 'Revenue sensitivity' },
    ])

    engine.insertColumns('Labels', 0, 1)
    expect(richTextCells(engine.exportSnapshot())).toEqual([
      { address: 'B2', text: 'Important: Before signing off' },
      { address: 'C3', text: 'Revenue sensitivity' },
    ])

    engine.deleteRows('Labels', 2, 1)
    expect(richTextCells(engine.exportSnapshot())).toEqual([{ address: 'B2', text: 'Important: Before signing off' }])
  })
})

function richTextCells(snapshot: WorkbookSnapshot): Array<{ readonly address: string; readonly text: string }> {
  return (
    snapshot.sheets[0]?.metadata?.richTextArtifacts?.cells.map((cell) => ({
      address: cell.address,
      text: cell.text,
    })) ?? []
  )
}

function richTextSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Rich text runtime' },
    sheets: [
      {
        id: 1,
        name: 'Labels',
        order: 0,
        metadata: { richTextArtifacts },
        cells: [
          { address: 'A1', value: 'Important: Before signing off' },
          { address: 'B2', value: 'Revenue sensitivity' },
        ],
      },
    ],
  }
}
