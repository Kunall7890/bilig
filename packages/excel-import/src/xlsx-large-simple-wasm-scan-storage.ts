import { createWorksheetImportStorageSync, type WorksheetImportStorage } from '@bilig/wasm-kernel'

import type { LiteralInput } from '@bilig/protocol'
import type { LargeSimpleFormulaNumericRecords } from './xlsx-large-simple-formula-records.js'
import type { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-arena.js'

export type LargeSimpleWorksheetScanStorageKind = 'js' | 'wasm' | 'wasm-fallback'

export class LargeSimpleWorksheetScanStorageBridge {
  private wasmStorage: LargeSimpleWorksheetWasmScanStorage | null
  private scanStorageKind: LargeSimpleWorksheetScanStorageKind

  constructor(
    private readonly sheetIndex: number,
    private readonly arena: ImportedWorkbookArena,
    private readonly styleIndexes: ImportedWorksheetStyleIndexArena,
    private readonly options: {
      readonly useWasmScanStorage?: boolean
      readonly retainCells: boolean
      readonly onWasmFormulaRecordsBeforeRelease?: (records: LargeSimpleFormulaNumericRecords) => void
    },
  ) {
    this.wasmStorage = options.useWasmScanStorage === false || !options.retainCells ? null : LargeSimpleWorksheetWasmScanStorage.create()
    this.scanStorageKind = this.wasmStorage ? 'wasm' : 'js'
  }

  get kind(): LargeSimpleWorksheetScanStorageKind {
    return this.scanStorageKind
  }

  addCell(row: number, column: number, value: LiteralInput | undefined): number {
    if (this.wasmStorage) {
      if (typeof value === 'number') {
        return this.wasmStorage.addNumberCell(row, column, value)
      }
      if (value === undefined) {
        return this.wasmStorage.addFormulaOnlyCell(row, column)
      }
      this.flushForFallback()
    }
    return this.arena.addCell({ sheetIndex: this.sheetIndex, row, column, value })
  }

  addSharedStringCell(row: number, column: number, sharedStringIndex: number): number {
    if (this.wasmStorage) {
      return this.wasmStorage.addSharedStringCell(row, column, sharedStringIndex)
    }
    return this.arena.addSharedStringCell({ sheetIndex: this.sheetIndex, row, column, sharedStringIndex })
  }

  addStyle(row: number, column: number, styleIndex: number): void {
    if (this.wasmStorage) {
      this.wasmStorage.addStyle(row, column, styleIndex)
      return
    }
    this.styleIndexes.add(row, column, styleIndex)
  }

  addFormulaRecord(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null): boolean {
    if (!this.wasmStorage) {
      return false
    }
    this.wasmStorage.addFormulaRecord(cellIndex, row, column, typeCode, sharedIndex)
    return true
  }

  flushCellsAndStyles(): void {
    if (!this.wasmStorage) {
      return
    }
    this.wasmStorage.transferCellsAndStylesToArena(this.sheetIndex, this.arena, this.styleIndexes)
  }

  formulaRecords(): LargeSimpleFormulaNumericRecords | undefined {
    return this.wasmStorage?.formulaRecords()
  }

  flush(): void {
    if (!this.wasmStorage) {
      return
    }
    this.flushCellsAndStyles()
    this.release()
  }

  release(): void {
    if (!this.wasmStorage) {
      return
    }
    this.wasmStorage.release()
    this.wasmStorage = null
  }

  private flushForFallback(): void {
    if (this.wasmStorage && this.scanStorageKind === 'wasm') {
      this.scanStorageKind = 'wasm-fallback'
      const formulaRecords = this.wasmStorage.formulaRecords()
      if (formulaRecords.cellIndexes.length > 0) {
        this.options.onWasmFormulaRecordsBeforeRelease?.(formulaRecords)
      }
    }
    this.flush()
  }
}

class LargeSimpleWorksheetWasmScanStorage {
  private constructor(private readonly storage: WorksheetImportStorage) {}

  static create(): LargeSimpleWorksheetWasmScanStorage | null {
    if (!isNodeLike()) {
      return null
    }
    try {
      const storage = createWorksheetImportStorageSync()
      storage.reset(64, 16, 16)
      return new LargeSimpleWorksheetWasmScanStorage(storage)
    } catch {
      return null
    }
  }

  addNumberCell(row: number, column: number, value: number): number {
    return this.storage.addNumberCell(row, column, value)
  }

  addFormulaOnlyCell(row: number, column: number): number {
    return this.storage.addFormulaOnlyCell(row, column)
  }

  addSharedStringCell(row: number, column: number, sharedStringIndex: number): number {
    return this.storage.addSharedStringCell(row, column, sharedStringIndex)
  }

  addStyle(row: number, column: number, styleIndex: number): void {
    this.storage.addStyle(row, column, styleIndex)
  }

  addFormulaRecord(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null): void {
    this.storage.addFormulaRecord(cellIndex, row, column, typeCode, sharedIndex)
  }

  formulaRecords(): LargeSimpleFormulaNumericRecords {
    const snapshot = this.storage.snapshot()
    return {
      cellIndexes: snapshot.formulaCellIndexes,
      rows: snapshot.formulaRows,
      columns: snapshot.formulaColumns,
      typeCodes: snapshot.formulaTypeCodes,
      sharedIndexes: snapshot.formulaSharedIndexes,
    }
  }

  transferCellsAndStylesToArena(sheetIndex: number, arena: ImportedWorkbookArena, styleIndexes: ImportedWorksheetStyleIndexArena): void {
    const snapshot = this.storage.snapshot()
    for (let index = 0; index < snapshot.rows.length; index += 1) {
      const row = snapshot.rows[index] ?? 0
      const column = snapshot.columns[index] ?? 0
      const valueKind = snapshot.valueKinds[index] ?? this.storage.valueKindFormulaOnly
      switch (valueKind) {
        case this.storage.valueKindNumber:
          arena.addCell({ sheetIndex, row, column, value: snapshot.numbers[index] })
          break
        case this.storage.valueKindSharedString:
          arena.addSharedStringCell({ sheetIndex, row, column, sharedStringIndex: snapshot.sharedStringIds[index] ?? 0 })
          break
        default:
          arena.addCell({ sheetIndex, row, column, value: undefined })
          break
      }
    }
    for (let index = 0; index < snapshot.styleRows.length; index += 1) {
      styleIndexes.add(snapshot.styleRows[index] ?? 0, snapshot.styleColumns[index] ?? 0, snapshot.styleIds[index] ?? 0)
    }
  }

  release(): void {
    this.storage.release()
  }
}

function isNodeLike(): boolean {
  return typeof process !== 'undefined' && process.versions != null && process.versions.node != null
}
