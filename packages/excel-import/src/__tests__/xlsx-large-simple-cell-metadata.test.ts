import type { WorkbookSnapshot } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena } from '../xlsx-large-simple-arena.js'
import {
  buildLargeSimpleCellMetadataReferenceSnapshots,
  buildLargeSimpleLazyCellMetadataReferenceSnapshots,
} from '../xlsx-large-simple-cell-metadata.js'

type WorkbookSheetCells = WorkbookSnapshot['sheets'][number]['cells']

describe('large simple XLSX cell metadata', () => {
  it('signs lazy cell metadata refs from the arena without expanding the public cell list', () => {
    const arena = new ImportedWorkbookArena()
    arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 'unused' })
    const formulaCell = arena.addCell({ sheetIndex: 0, row: 1, column: 1, value: 9 })
    arena.setFormula(formulaCell, 'A1+1')

    const emptyCells: WorkbookSheetCells = []
    const cells = new Proxy(emptyCells, {
      get(target, property, receiver) {
        if (property === 'map') {
          throw new Error('lazy cells should not be expanded for metadata refs')
        }
        return Reflect.get(target, property, receiver)
      },
    })

    expect(
      buildLargeSimpleCellMetadataReferenceSnapshots(
        [
          { address: 'B2', vm: '4' },
          { address: 'Z99', cm: '1' },
        ],
        cells,
        { arena, sheetIndex: 0 },
        true,
      ),
    ).toEqual([
      {
        address: 'B2',
        vm: '4',
        cellSignature: JSON.stringify({ value: 9, formula: 'A1+1', format: null }),
      },
      {
        address: 'Z99',
        cm: '1',
        cellSignature: 'null',
      },
    ])
  })

  it('signs lazy metadata before detached pool transfer releases arena strings and formulas', () => {
    const arena = new ImportedWorkbookArena()
    const formulaCell = arena.addCell({ sheetIndex: 0, row: 1, column: 1, value: 'Cached label' })
    arena.setFormula(formulaCell, 'A1&"!"')

    const metadataRefs = buildLargeSimpleLazyCellMetadataReferenceSnapshots([{ address: 'B2', vm: '9' }], {
      arena,
      sheetIndex: 0,
    })
    const detachedCells = arena.createDetachedLazySheetCells(0)
    arena.release()

    expect(metadataRefs).toEqual([
      {
        address: 'B2',
        vm: '9',
        cellSignature: JSON.stringify({ value: 'Cached label', formula: 'A1&"!"', format: null }),
      },
    ])
    expect(detachedCells[0]).toEqual({ address: 'B2', value: 'Cached label', formula: 'A1&"!"' })
  })
})
