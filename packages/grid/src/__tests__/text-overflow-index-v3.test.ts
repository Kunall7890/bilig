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

  test('clears inbound spill records when the replaced tile no longer contains the spill run', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 120, spillColEnd: 135 })

    index.replaceTileRuns({
      sheetOrdinal: 2,
      textRuns: [],
      viewport: { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
    })

    const dependencies: string[] = []
    index.markDependenciesForCell({ sheetOrdinal: 2, row: 4, col: 130 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual([])
  })

  test('keeps inbound spill records when the replaced tile still carries the spill run', () => {
    const index = new TextOverflowIndexV3()
    index.updateSpill({ sheetOrdinal: 2, row: 4, col: 120, spillColEnd: 135 })

    index.replaceTileRuns({
      sheetOrdinal: 2,
      textRuns: [
        {
          clipHeight: 20,
          clipWidth: 200,
          clipX: 0,
          clipY: 0,
          col: 120,
          color: '#111827',
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 4,
          spillColEnd: 135,
          strike: false,
          text: 'spills into current tile',
          underline: false,
          width: 600,
          x: -800,
          y: 88,
        },
      ],
      viewport: { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
    })

    const dependencies: string[] = []
    index.markDependenciesForCell({ sheetOrdinal: 2, row: 4, col: 130 }, (source) => {
      dependencies.push(`${source.row}:${source.col}:${source.spillColEnd}`)
    })

    expect(dependencies).toEqual(['4:120:135'])
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
