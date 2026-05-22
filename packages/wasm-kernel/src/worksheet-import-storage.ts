import type { RawKernelExports } from './raw-kernel-exports.js'

export interface WorksheetImportStorageSnapshot {
  readonly rows: Uint32Array
  readonly columns: Uint16Array
  readonly valueKinds: Uint8Array
  readonly numbers: Float64Array
  readonly sharedStringIds: Uint32Array
  readonly styleRows: Uint32Array
  readonly styleColumns: Uint16Array
  readonly styleIds: Uint32Array
  readonly formulaCellIndexes: Uint32Array
  readonly formulaRows: Uint32Array
  readonly formulaColumns: Uint16Array
  readonly formulaTypeCodes: Uint8Array
  readonly formulaSharedIndexes: Uint32Array
}

export interface WorksheetImportStorage {
  readonly valueKindFormulaOnly: number
  readonly valueKindNumber: number
  readonly valueKindSharedString: number
  readonly noSharedFormulaIndex: number
  reset(cellCapacity: number, styleCapacity: number, formulaCapacity: number): void
  release(): void
  addNumberCellFromBytes(row: number, column: number, bytes: Uint8Array, start: number, end: number): number | null
  readNonNegativeIntegerFromBytes(bytes: Uint8Array, start: number, end: number): number | null
  addNumberCell(row: number, column: number, value: number): number
  addFormulaOnlyCell(row: number, column: number): number
  addSharedStringCell(row: number, column: number, sharedStringIndex: number): number
  addStyle(row: number, column: number, styleId: number): void
  addFormulaRecord(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null): void
  snapshot(): WorksheetImportStorageSnapshot
}

export class RawWorksheetImportStorage implements WorksheetImportStorage {
  readonly valueKindFormulaOnly: number
  readonly valueKindNumber: number
  readonly valueKindSharedString: number
  readonly noSharedFormulaIndex: number

  constructor(private readonly raw: RawKernelExports) {
    this.valueKindFormulaOnly = raw.worksheetImportValueKindFormulaOnly.value
    this.valueKindNumber = raw.worksheetImportValueKindNumber.value
    this.valueKindSharedString = raw.worksheetImportValueKindSharedString.value
    this.noSharedFormulaIndex = raw.worksheetImportNoSharedFormulaIndex.value >>> 0
  }

  reset(cellCapacity: number, styleCapacity: number, formulaCapacity: number): void {
    this.raw.resetWorksheetImportStorage(cellCapacity, styleCapacity, formulaCapacity)
  }

  release(): void {
    this.raw.releaseWorksheetImportStorage()
  }

  addNumberCellFromBytes(row: number, column: number, bytes: Uint8Array, start: number, end: number): number | null {
    const length = this.copyScratch(bytes, start, end)
    const cellIndex = this.raw.addWorksheetImportNumberCellFromScratch(row, column, length)
    return cellIndex >= 0 ? cellIndex : null
  }

  readNonNegativeIntegerFromBytes(bytes: Uint8Array, start: number, end: number): number | null {
    const length = this.copyScratch(bytes, start, end)
    const value = this.raw.readWorksheetImportNonNegativeIntegerFromScratch(length)
    return Number.isSafeInteger(value) && value >= 0 ? value : null
  }

  addNumberCell(row: number, column: number, value: number): number {
    return this.raw.addWorksheetImportNumberCell(row, column, value)
  }

  addFormulaOnlyCell(row: number, column: number): number {
    return this.raw.addWorksheetImportFormulaOnlyCell(row, column)
  }

  addSharedStringCell(row: number, column: number, sharedStringIndex: number): number {
    return this.raw.addWorksheetImportSharedStringCell(row, column, sharedStringIndex)
  }

  addStyle(row: number, column: number, styleId: number): void {
    this.raw.addWorksheetImportStyle(row, column, styleId)
  }

  addFormulaRecord(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null): void {
    this.raw.addWorksheetImportFormulaRecord(
      cellIndex,
      row,
      column,
      typeCode,
      sharedIndex === null ? this.noSharedFormulaIndex : sharedIndex,
    )
  }

  snapshot(): WorksheetImportStorageSnapshot {
    const memory = this.raw.memory.buffer
    const cellCount = this.raw.getWorksheetImportCellCount()
    const styleCount = this.raw.getWorksheetImportStyleCount()
    const formulaCount = this.raw.getWorksheetImportFormulaCount()
    return {
      rows: new Uint32Array(memory, this.raw.getWorksheetImportRowsPtr(), cellCount),
      columns: new Uint16Array(memory, this.raw.getWorksheetImportColumnsPtr(), cellCount),
      valueKinds: new Uint8Array(memory, this.raw.getWorksheetImportValueKindsPtr(), cellCount),
      numbers: new Float64Array(memory, this.raw.getWorksheetImportNumbersPtr(), cellCount),
      sharedStringIds: new Uint32Array(memory, this.raw.getWorksheetImportSharedStringIdsPtr(), cellCount),
      styleRows: new Uint32Array(memory, this.raw.getWorksheetImportStyleRowsPtr(), styleCount),
      styleColumns: new Uint16Array(memory, this.raw.getWorksheetImportStyleColumnsPtr(), styleCount),
      styleIds: new Uint32Array(memory, this.raw.getWorksheetImportStyleIdsPtr(), styleCount),
      formulaCellIndexes: new Uint32Array(memory, this.raw.getWorksheetImportFormulaCellIndexesPtr(), formulaCount),
      formulaRows: new Uint32Array(memory, this.raw.getWorksheetImportFormulaRowsPtr(), formulaCount),
      formulaColumns: new Uint16Array(memory, this.raw.getWorksheetImportFormulaColumnsPtr(), formulaCount),
      formulaTypeCodes: new Uint8Array(memory, this.raw.getWorksheetImportFormulaTypeCodesPtr(), formulaCount),
      formulaSharedIndexes: new Uint32Array(memory, this.raw.getWorksheetImportFormulaSharedIndexesPtr(), formulaCount),
    }
  }

  private copyScratch(bytes: Uint8Array, start: number, end: number): number {
    const safeStart = Math.max(0, Math.min(bytes.byteLength, Math.trunc(start)))
    const safeEnd = Math.max(safeStart, Math.min(bytes.byteLength, Math.trunc(end)))
    const length = safeEnd - safeStart
    const scratchPtr = this.raw.prepareWorksheetImportScratch(length)
    new Uint8Array(this.raw.memory.buffer, scratchPtr, length).set(bytes.subarray(safeStart, safeEnd))
    return length
  }
}
