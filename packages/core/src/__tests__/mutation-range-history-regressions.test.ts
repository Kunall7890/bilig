import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('range mutation history regressions', () => {
  it('restores inherited number formats when undoing a blank copy over a formatted cell', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'range-copy-undo-format-regression' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Region')
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, '0.00')

    const formattedSnapshot = engine.exportSnapshot()

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'E1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
    )
    expect(engine.undo()).toBe(true)

    expect(engine.exportSnapshot()).toEqual(formattedSnapshot)
  })
})
