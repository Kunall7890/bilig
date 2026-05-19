import { describe, expect, test } from 'vitest'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { createGridNavigationResolver } from '../gridNavigation.js'

function textCell(address: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    input: address,
    value: { tag: ValueTag.String, value: address, stringId: 0 },
    flags: 0,
    version: 1,
  }
}

function styledEmptyCell(address: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    styleId: 'style-green',
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 1,
  }
}

function emptyCell(address: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  }
}

function createEngine(cells: Record<string, CellSnapshot>): GridEngineLike {
  return {
    getCell: (_sheetName, address) => cells[address] ?? emptyCell(address),
    getCellStyle: () => undefined,
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => ({
        grid: {
          forEachCellEntry: (listener) => {
            Object.keys(cells).forEach((address, index) => {
              const parsed = parseCellAddress(address, 'Sheet1')
              listener(index, parsed.row, parsed.col)
            })
          },
        },
      }),
    },
  }
}

describe('gridNavigation', () => {
  test('moves primary-arrow navigation to contiguous data boundaries and then the next island', () => {
    const engine = createEngine({
      A1: textCell('A1'),
      A2: textCell('A2'),
      A3: textCell('A3'),
      A5: textCell('A5'),
      B2: textCell('B2'),
      C2: textCell('C2'),
      D2: textCell('D2'),
      F2: textCell('F2'),
    })
    const resolver = createGridNavigationResolver({ engine, sheetName: 'Sheet1' })

    expect(resolver.resolveDataEdge([0, 0], 'down')).toEqual([0, 2])
    expect(resolver.resolveDataEdge([0, 2], 'down')).toEqual([0, 4])
    expect(resolver.resolveDataEdge([0, 3], 'down')).toEqual([0, 4])
    expect(resolver.resolveDataEdge([1, 1], 'right')).toEqual([3, 1])
    expect(resolver.resolveDataEdge([3, 1], 'right')).toEqual([5, 1])
    expect(resolver.resolveDataEdge([4, 1], 'right')).toEqual([5, 1])
  })

  test('resolves the current data region while ignoring formatting-only empty cells', () => {
    const cells = Object.fromEntries(['B2', 'C2', 'D2', 'B3', 'D3', 'B4', 'C4', 'D4'].map((address) => [address, textCell(address)]))
    const engine = createEngine({
      ...cells,
      H8: styledEmptyCell('H8'),
    })
    const resolver = createGridNavigationResolver({ engine, sheetName: 'Sheet1' })

    expect(resolver.resolveCurrentRegion([2, 2])).toEqual({ x: 1, y: 1, width: 3, height: 3 })
    expect(resolver.resolveCurrentRegion([7, 7])).toBeNull()
  })

  test('falls back to sheet edges when no content exists in the requested direction', () => {
    const engine = createEngine({
      [formatAddress(3, 3)]: textCell('D4'),
    })
    const resolver = createGridNavigationResolver({ engine, sheetName: 'Sheet1' })

    expect(resolver.resolveDataEdge([3, 3], 'left')).toEqual([0, 3])
    expect(resolver.resolveDataEdge([3, 3], 'up')).toEqual([3, 0])
  })
})
