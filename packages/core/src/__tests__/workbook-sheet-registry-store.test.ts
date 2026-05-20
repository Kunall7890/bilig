import { describe, expect, it, vi } from 'vitest'
import { WorkbookSheetRegistryStore } from '../workbook-sheet-registry-store.js'
import { createWorkbookMetadataRecord } from '../workbook-metadata-types.js'
import type { SheetRecord } from '../workbook-sheet-record.js'

function createRegistry() {
  const sheetsByName = new Map<string, SheetRecord>()
  const sheetsById = new Map<number, SheetRecord>()
  const metadata = createWorkbookMetadataRecord()
  const deleteSheetRecords = vi.fn()
  const renameSheetRecords = vi.fn()
  const registry = new WorkbookSheetRegistryStore({
    sheetsByName,
    sheetsById,
    metadata,
    counters: undefined,
    cellKeyToIndex: new Map(),
    cellFormats: new Map(),
    getCellPosition: () => undefined,
    deleteSheetRecords,
    renameSheetRecords,
  })
  return { registry, sheetsByName, sheetsById, metadata, deleteSheetRecords, renameSheetRecords }
}

describe('WorkbookSheetRegistryStore', () => {
  it('creates, looks up, bumps, deletes, and resets sheet records', () => {
    const { registry, sheetsByName, sheetsById, deleteSheetRecords } = createRegistry()

    const first = registry.createSheet('Sheet1', 0, 5)
    expect(first.id).toBe(5)
    expect(registry.createSheet('Sheet1', 2, 9)).toBe(first)
    expect(first.id).toBe(9)
    expect(first.order).toBe(2)
    expect(sheetsById.has(5)).toBe(false)
    expect(sheetsById.get(9)).toBe(first)

    const second = registry.createSheet('Sheet2')
    expect(second.id).toBe(10)
    expect(registry.getSheet('Sheet1')).toBe(first)
    expect(registry.getSheetById(10)).toBe(second)
    expect(registry.getOrCreateSheet('Sheet3').id).toBe(11)
    expect(registry.getSheetNameById(10)).toBe('Sheet2')
    expect(registry.getSheetNameById(999)).toBe('')
    expect(registry.getSheetColumnVersion('Missing', 0)).toBe(0)
    expect(registry.getSheetStructureVersion('Missing')).toBe(0)

    registry.deleteSheet('Missing')
    registry.deleteSheet('Sheet3')
    expect(deleteSheetRecords).toHaveBeenCalledWith('Sheet3')
    expect(sheetsByName.has('Sheet3')).toBe(false)

    registry.reset()
    expect(sheetsByName.size).toBe(0)
    expect(sheetsById.size).toBe(0)
    expect(registry.createSheet('AfterReset').id).toBe(1)
  })

  it('renames by name and id while preserving sheet-scoped ranges', () => {
    const { registry, metadata, renameSheetRecords } = createRegistry()
    const sheet = registry.createSheet('Source', 0, 1)
    registry.createSheet('Taken', 1, 2)
    metadata.tables.set('table', {
      name: 'Table1',
      sheetName: 'Source',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['A', 'B'],
      headerRow: true,
      totalsRow: false,
    })
    sheet.styleRanges = [{ range: { sheetName: 'Source', startAddress: 'A1', endAddress: 'A1' }, styleId: 'style-1' }]
    sheet.formatRanges = [{ range: { sheetName: 'Source', startAddress: 'B1', endAddress: 'B1' }, formatId: 'format-1' }]

    expect(() => registry.renameSheet('Source', '   ')).toThrow('Sheet name must be non-empty')
    expect(registry.renameSheet('Missing', 'Other')).toBeUndefined()
    expect(registry.renameSheet('Source', 'Taken')).toBeUndefined()
    expect(registry.renameSheet('Source', 'Source')).toBe(sheet)

    expect(registry.renameSheet('Source', ' Renamed ')).toBe(sheet)
    expect(sheet.name).toBe('Renamed')
    expect(renameSheetRecords).toHaveBeenCalledWith('Source', 'Renamed')
    expect(sheet.styleRanges[0]?.range.sheetName).toBe('Renamed')
    expect(sheet.formatRanges[0]?.range.sheetName).toBe('Renamed')

    sheet.styleRanges = [{ range: { sheetName: 'Renamed', startAddress: 'C1', endAddress: 'C1' }, styleId: 'style-2' }]
    sheet.formatRanges = [{ range: { sheetName: 'Renamed', startAddress: 'D1', endAddress: 'D1' }, formatId: 'format-2' }]

    expect(() => registry.renameSheetById(sheet.id, '')).toThrow('Sheet name must be non-empty')
    expect(registry.renameSheetById(999, 'Ghost')).toBeUndefined()
    expect(registry.renameSheetById(sheet.id, 'Renamed')).toBe(sheet)
    expect(registry.renameSheetById(sheet.id, 'Taken')).toBeUndefined()

    expect(registry.renameSheetById(sheet.id, 'Final')).toBe(sheet)
    expect(sheet.name).toBe('Final')
    expect(renameSheetRecords).toHaveBeenCalledWith('Renamed', 'Final')
    expect(sheet.styleRanges[0]?.range.sheetName).toBe('Final')
    expect(sheet.formatRanges[0]?.range.sheetName).toBe('Final')
  })
})
