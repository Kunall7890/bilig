import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

const ignoredErrorsXml =
  '<ignoredErrors><ignoredError sqref="B2:B4 D5" numberStoredAsText="1"/><ignoredError sqref="C3" formula="1"/></ignoredErrors>'

describe('SpreadsheetEngine ignored error metadata', () => {
  it('roundtrips worksheet ignoredErrors XML through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'ignored-errors-runtime' })
    await engine.ready()

    engine.importSnapshot(ignoredErrorsSnapshot())
    engine.setCellValue('Review', 'A6', 'agent edit')

    const exported = engine.exportSnapshot()
    expect(exported.sheets[0]?.metadata?.ignoredErrors).toEqual({ xml: ignoredErrorsXml })

    const restored = new SpreadsheetEngine({ workbookName: 'ignored-errors-runtime-restored' })
    await restored.ready()
    restored.importSnapshot(exported)

    expect(restored.exportSnapshot().sheets[0]?.metadata?.ignoredErrors).toEqual({ xml: ignoredErrorsXml })
  })

  it('rewrites ignored error sqref ranges through structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'ignored-errors-structural' })
    await engine.ready()
    engine.importSnapshot(ignoredErrorsSnapshot())

    engine.insertRows('Review', 0, 1)

    expect(ignoredErrorSqrefs(engine.exportSnapshot().sheets[0]?.metadata?.ignoredErrors?.xml)).toEqual(['B3:B5 D6', 'C4'])
  })
})

function ignoredErrorSqrefs(xml: string | undefined): string[] {
  if (!xml) {
    throw new Error('Expected ignoredErrors XML')
  }
  return [...xml.matchAll(/\bignoredError\b[^>]*\bsqref="([^"]*)"/gu)].map((match) => match[1] ?? '')
}

function ignoredErrorsSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Ignored errors' },
    sheets: [
      {
        id: 1,
        name: 'Review',
        order: 0,
        metadata: {
          ignoredErrors: { xml: ignoredErrorsXml },
        },
        cells: [
          { address: 'A1', value: 'Header' },
          { address: 'B2', value: '001' },
          { address: 'B3', value: '002' },
          { address: 'B4', value: '003' },
          { address: 'C3', formula: 'A1', value: 'Header' },
          { address: 'D5', value: '004' },
        ],
      },
    ],
  }
}
