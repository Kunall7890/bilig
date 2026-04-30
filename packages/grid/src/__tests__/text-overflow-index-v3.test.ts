import { describe, expect, test } from 'vitest'
import { TextOverflowIndexV3 } from '../renderer-v3/text-overflow-index.js'

describe('TextOverflowIndexV3', () => {
  test('marks source spill ranges when a blocker cell changes', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 1, spillColEnd: 6 })

    const dependencies: string[] = []
    index.markDependenciesForCell({ sheetOrdinal: 2, row: 4, col: 3 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual(['4:1:6'])
  })

  test('marks source spill ranges when the source cell changes', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 1, spillColEnd: 6 })

    const dependencies: string[] = []
    index.markDependenciesForCell({ sheetOrdinal: 2, row: 4, col: 1 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual(['4:1:6'])
  })

  test('replaces tile-owned source records when a tile is rematerialized', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 1, spillColEnd: 6 })

    index.replaceTileRuns({
      sheetOrdinal: 2,
      textRuns: [],
      viewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })

    const dependencies: string[] = []
    index.markDependenciesForCell({ sheetOrdinal: 2, row: 4, col: 3 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual([])
  })

  test('marks column dependencies for resize invalidation without unrelated rows', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 1, spillColEnd: 6 })
    index.updateSpill({ sheetOrdinal: 2, row: 8, col: 9, spillColEnd: 12 })

    const dependencies: string[] = []
    index.markDependenciesForAxisX({ sheetOrdinal: 2, colStart: 3, colEnd: 3 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual(['4:1:6'])
  })
})
