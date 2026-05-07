import { describe, expect, it } from 'vitest'
import { readRuntimeSnapshot } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  buildWorkPaperDenseRange,
  collectSerializedWorkPaperSheets,
  collectWorkPaperSheetsByName,
  readWorkPaperSheetRange,
  workPaperRangeForSheetDimensions,
} from '../work-paper-sheet-read.js'

describe('work paper sheet read helpers', () => {
  it('builds dense ranges by visiting each address', () => {
    expect(
      buildWorkPaperDenseRange(
        {
          start: { sheet: 2, row: 3, col: 4 },
          end: { sheet: 2, row: 4, col: 6 },
        },
        (address) => `${address.sheet}:${address.row}:${address.col}`,
      ),
    ).toEqual([
      ['2:3:4', '2:3:5', '2:3:6'],
      ['2:4:4', '2:4:5', '2:4:6'],
    ])
  })

  it('maps non-empty sheet dimensions to a full-sheet range', () => {
    expect(workPaperRangeForSheetDimensions(3, { width: 2, height: 4 })).toEqual({
      start: { sheet: 3, row: 0, col: 0 },
      end: { sheet: 3, row: 3, col: 1 },
    })
    expect(workPaperRangeForSheetDimensions(3, { width: 0, height: 4 })).toBeUndefined()
    expect(readWorkPaperSheetRange({ sheetId: 3, dimensions: { width: 0, height: 0 }, readRange: () => [[1]] })).toEqual([])
  })

  it('collects sheet records by name and preserves runtime snapshots on serialized output', () => {
    const sheets = [
      { id: 1, name: 'Sheet1' },
      { id: 2, name: 'Sheet2' },
    ]
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Book' },
      sheets: [],
    }

    expect(collectWorkPaperSheetsByName(sheets, (sheet) => sheet.id)).toEqual({ Sheet1: 1, Sheet2: 2 })
    const serialized = collectSerializedWorkPaperSheets({
      sheets,
      readSheet: (sheet) => [[sheet.name]],
      runtimeSnapshot: snapshot,
    })

    expect(serialized).toEqual({ Sheet1: [['Sheet1']], Sheet2: [['Sheet2']] })
    expect(readRuntimeSnapshot(serialized)).toBe(snapshot)
  })
})
