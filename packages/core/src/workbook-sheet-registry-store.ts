import type { EngineCounters } from './perf/engine-counters.js'
import { makeCellKey } from './workbook-cell-key-index.js'
import type { WorkbookMetadataRecord } from './workbook-metadata-types.js'
import { createWorkbookSheetRecord, type SheetRecord } from './workbook-sheet-record.js'

export class WorkbookSheetRegistryStore {
  private nextSheetId = 1

  constructor(
    private readonly options: {
      readonly sheetsByName: Map<string, SheetRecord>
      readonly sheetsById: Map<number, SheetRecord>
      readonly metadata: WorkbookMetadataRecord
      readonly counters: EngineCounters | undefined
      readonly cellKeyToIndex: Map<number, number>
      readonly cellFormats: Map<number, string>
      readonly getCellPosition: (cellIndex: number) => { sheetId: number; row: number; col: number } | undefined
      readonly deleteSheetRecords: (sheetName: string) => void
      readonly renameSheetRecords: (oldName: string, nextName: string) => void
    },
  ) {}

  createSheet(name: string, order = this.options.sheetsByName.size, id?: number): SheetRecord {
    const existing = this.options.sheetsByName.get(name)
    if (existing) {
      existing.order = order
      if (id !== undefined && existing.id !== id) {
        this.options.sheetsById.delete(existing.id)
        existing.id = id
        existing.logical.setSheetId(id)
        this.options.sheetsById.set(existing.id, existing)
        this.bumpSheetId(id)
      }
      return existing
    }
    const sheetId = id ?? this.nextSheetId++
    const sheet = createWorkbookSheetRecord({
      id: sheetId,
      name,
      order,
      counters: this.options.counters,
    })
    if (id !== undefined) {
      this.bumpSheetId(id)
    }
    this.options.sheetsByName.set(name, sheet)
    this.options.sheetsById.set(sheet.id, sheet)
    return sheet
  }

  deleteSheet(name: string): void {
    const sheet = this.options.sheetsByName.get(name)
    if (!sheet) {
      return
    }
    sheet.grid.forEachCell((cellIndex) => {
      const position = this.options.getCellPosition(cellIndex)
      if (position) {
        this.options.cellKeyToIndex.delete(makeCellKey(sheet.id, position.row, position.col))
      }
      this.options.cellFormats.delete(cellIndex)
      const identity = sheet.logical.getCellIdentity(cellIndex)
      if (identity) {
        sheet.logical.deleteVisibleCellByIds(identity.rowId, identity.colId)
      }
    })
    this.options.deleteSheetRecords(name)
    sheet.rowAxis.length = 0
    sheet.columnAxis.length = 0
    sheet.styleRanges.length = 0
    sheet.formatRanges.length = 0
    delete sheet.arrayFormulas
    delete sheet.dataTableFormulas
    delete sheet.sheetPr
    delete sheet.ignoredErrors
    delete sheet.printPageSetup
    delete sheet.printerSettings
    delete sheet.sparklines
    delete sheet.controlArtifacts
    delete sheet.richTextArtifacts
    delete sheet.cellMetadataRefs
    this.options.sheetsByName.delete(name)
    this.options.sheetsById.delete(sheet.id)
  }

  renameSheet(oldName: string, nextName: string): SheetRecord | undefined {
    const trimmedName = nextName.trim()
    if (trimmedName.length === 0) {
      throw new Error('Sheet name must be non-empty')
    }
    const sheet = this.options.sheetsByName.get(oldName)
    if (!sheet) {
      return undefined
    }
    if (oldName === trimmedName) {
      return sheet
    }
    if (this.options.sheetsByName.has(trimmedName)) {
      return undefined
    }

    this.options.sheetsByName.delete(oldName)
    sheet.name = trimmedName
    this.options.sheetsByName.set(trimmedName, sheet)
    this.options.renameSheetRecords(oldName, trimmedName)

    if (sheet.styleRanges.length > 0) {
      sheet.styleRanges = sheet.styleRanges.map((record) =>
        record.range.sheetName === oldName ? { ...record, range: { ...record.range, sheetName: trimmedName } } : record,
      )
    }
    if (sheet.formatRanges.length > 0) {
      sheet.formatRanges = sheet.formatRanges.map((record) =>
        record.range.sheetName === oldName ? { ...record, range: { ...record.range, sheetName: trimmedName } } : record,
      )
    }

    return sheet
  }

  renameSheetById(sheetId: number, trimmedName: string): SheetRecord | undefined {
    if (trimmedName.length === 0) {
      throw new Error('Sheet name must be non-empty')
    }
    const sheet = this.options.sheetsById.get(sheetId)
    if (!sheet) {
      return undefined
    }
    const oldName = sheet.name
    if (oldName === trimmedName) {
      return sheet
    }
    if (this.options.sheetsByName.has(trimmedName)) {
      return undefined
    }

    this.options.sheetsByName.delete(oldName)
    sheet.name = trimmedName
    this.options.sheetsByName.set(trimmedName, sheet)
    this.options.renameSheetRecords(oldName, trimmedName)

    if (sheet.styleRanges.length > 0) {
      sheet.styleRanges = sheet.styleRanges.map((record) =>
        record.range.sheetName === oldName ? { ...record, range: { ...record.range, sheetName: trimmedName } } : record,
      )
    }
    if (sheet.formatRanges.length > 0) {
      sheet.formatRanges = sheet.formatRanges.map((record) =>
        record.range.sheetName === oldName ? { ...record, range: { ...record.range, sheetName: trimmedName } } : record,
      )
    }

    return sheet
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.options.sheetsByName.get(name)
  }

  getSheetById(id: number): SheetRecord | undefined {
    return this.options.sheetsById.get(id)
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name)
  }

  getSheetNameById(id: number): string {
    return this.options.sheetsById.get(id)?.name ?? ''
  }

  getSheetColumnVersion(sheetName: string, col: number): number {
    return this.options.sheetsByName.get(sheetName)?.columnVersions[col] ?? 0
  }

  getSheetStructureVersion(sheetName: string): number {
    return this.options.sheetsByName.get(sheetName)?.structureVersion ?? 0
  }

  reset(): void {
    this.options.sheetsByName.clear()
    this.options.sheetsById.clear()
    this.nextSheetId = 1
  }

  private bumpSheetId(id: number): void {
    if (Number.isInteger(id) && id >= this.nextSheetId) {
      this.nextSheetId = id + 1
    }
  }
}
