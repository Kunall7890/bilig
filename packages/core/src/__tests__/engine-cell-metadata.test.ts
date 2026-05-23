import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const workbookCellMetadata = {
  relationshipTarget: 'metadata.xml',
  metadataXml: [
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" rowColShift="1" cellMeta="1"/></metadataTypes>',
    '<futureMetadata name="XLDAPR" count="1"><bk/></futureMetadata>',
    '<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>',
    '</metadata>',
  ].join(''),
}

const cellMetadataRefs = [
  {
    address: 'A2',
    cm: '1',
    cellSignature: cellSignature({ value: 'MSFT' }),
  },
  {
    address: 'B2',
    cm: '1',
    cellSignature: cellSignature({ value: 415.32 }),
  },
]

describe('engine cell metadata', () => {
  it('roundtrips workbook metadata parts and worksheet cell metadata refs through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-metadata-runtime' })
    await engine.ready()

    engine.importSnapshot(cellMetadataSnapshot())

    expect(engine.exportSnapshot().workbook.metadata?.cellMetadata).toEqual(workbookCellMetadata)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.cellMetadataRefs).toEqual(cellMetadataRefs)

    const restored = new SpreadsheetEngine({ workbookName: 'cell-metadata-restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.exportSnapshot().workbook.metadata?.cellMetadata).toEqual(workbookCellMetadata)
    expect(restored.exportSnapshot().sheets[0]?.metadata?.cellMetadataRefs).toEqual(cellMetadataRefs)
  })

  it('rewrites worksheet cell metadata refs during structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-metadata-structure' })
    await engine.ready()
    engine.importSnapshot(cellMetadataSnapshot())

    engine.insertRows('Rich Values', 0, 1)
    expect(metadataRefAddresses(engine.exportSnapshot())).toEqual(['A3', 'B3'])

    engine.insertColumns('Rich Values', 0, 1)
    expect(metadataRefAddresses(engine.exportSnapshot())).toEqual(['B3', 'C3'])

    engine.deleteRows('Rich Values', 2, 1)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.cellMetadataRefs).toBeUndefined()
    expect(engine.exportSnapshot().workbook.metadata?.cellMetadata).toEqual(workbookCellMetadata)
  })
})

function metadataRefAddresses(snapshot: WorkbookSnapshot): string[] {
  return snapshot.sheets[0]?.metadata?.cellMetadataRefs?.map((ref) => ref.address) ?? []
}

function cellMetadataSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Cell metadata runtime',
      metadata: { cellMetadata: workbookCellMetadata },
    },
    sheets: [
      {
        id: 1,
        name: 'Rich Values',
        order: 0,
        metadata: { cellMetadataRefs },
        cells: [
          { address: 'A2', value: 'MSFT' },
          { address: 'B2', value: 415.32 },
        ],
      },
    ],
  }
}

function cellSignature(cell: { readonly value: string | number }): string {
  return JSON.stringify({
    value: cell.value,
    formula: null,
    format: null,
  })
}
