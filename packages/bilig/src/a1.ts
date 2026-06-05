import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'
import type {
  PersistedWorkPaperDocument,
  RawCellContent,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperSheets,
} from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

export type A1CellInput = RawCellContent
export type A1CellValue = ReturnType<WorkPaperInstance['getCellValue']>

export interface A1WorkPaperOptions {
  readonly defaultSheetName?: string
  readonly writableSheets?: readonly string[]
  readonly limitations?: readonly string[]
}

export interface A1CellRead {
  readonly address: string
  readonly value: A1CellValue
  readonly serialized: RawCellContent
  readonly formula: string | null
  readonly displayValue: string
}

export interface A1RangeRead {
  readonly range: string
  readonly values: A1CellValue[][]
  readonly serialized: RawCellContent[][]
  readonly displayValues: string[][]
}

export interface A1SetCellAndReadbackOptions {
  readonly readbackRange: string
  readonly requireReadbackChange?: boolean
  readonly includeConfig?: boolean
  readonly includeSerializedDocument?: boolean
  readonly limitations?: readonly string[]
}

export interface A1SetCellAndReadbackProof {
  readonly editedCell: string
  readonly readbackRange: string
  readonly before: A1CellRead
  readonly after: A1CellRead
  readonly beforeReadback: A1RangeRead
  readonly afterReadback: A1RangeRead
  readonly restoredReadback: A1RangeRead
  readonly persistedDocumentBytes: number
  readonly checks: {
    readonly readbackChanged: boolean
    readonly restoredReadbackMatchesAfter: boolean
    readonly previousSerialized: RawCellContent
    readonly newSerialized: RawCellContent
  }
  readonly verified: boolean
  readonly limitations: readonly string[]
  readonly changes: readonly WorkPaperChange[]
  readonly serializedDocument?: string
}

export interface A1WorkPaper {
  readonly workpaper: WorkPaperInstance
  readCell(address: string): A1CellRead
  get(address: string): A1CellValue
  display(address: string): string
  formula(address: string): string | null
  readRange(range: string): A1RangeRead
  range(range: string): A1CellValue[][]
  setCell(address: string, value: A1CellInput): WorkPaperChange[]
  set(address: string, value: A1CellInput): WorkPaperChange[]
  readMany(addresses: readonly string[]): Readonly<Record<string, A1CellRead>>
  setCells(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[]
  setMany(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[]
  setCellAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof
  editAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof
  exportDocument(options?: { readonly includeConfig?: boolean }): PersistedWorkPaperDocument
  serialize(options?: { readonly includeConfig?: boolean }): string
  saveJson(options?: { readonly includeConfig?: boolean }): string
  restoreJson(input: string | PersistedWorkPaperDocument): A1WorkPaper
  dispose(): void
}

export function createA1WorkPaper(workpaper: WorkPaperInstance, options: A1WorkPaperOptions = {}): A1WorkPaper {
  return new A1WorkPaperFacade(workpaper, options)
}

export function buildA1WorkPaper(sheets: WorkPaperSheets, config?: WorkPaperConfig, options: A1WorkPaperOptions = {}): A1WorkPaper {
  return createA1WorkPaper(WorkPaper.buildFromSheets(sheets, config), options)
}

export function restoreA1WorkPaper(input: string | PersistedWorkPaperDocument, options: A1WorkPaperOptions = {}): A1WorkPaper {
  const document = typeof input === 'string' ? parseWorkPaperDocument(input) : input
  return createA1WorkPaper(createWorkPaperFromDocument(document), options)
}

class A1WorkPaperFacade implements A1WorkPaper {
  readonly workpaper: WorkPaperInstance
  private readonly defaultSheetName: string | undefined
  private readonly writableSheets: readonly string[] | undefined
  private readonly limitations: readonly string[]

  constructor(workpaper: WorkPaperInstance, options: A1WorkPaperOptions) {
    this.workpaper = workpaper
    this.defaultSheetName = options.defaultSheetName
    this.writableSheets = options.writableSheets
    this.limitations = options.limitations ?? []
  }

  readCell(address: string): A1CellRead {
    const parsed = this.requireCellAddress(address)
    return this.readParsedCell(parsed)
  }

  get(address: string): A1CellValue {
    return this.readCell(address).value
  }

  display(address: string): string {
    return this.readCell(address).displayValue
  }

  formula(address: string): string | null {
    return this.readCell(address).formula
  }

  readRange(range: string): A1RangeRead {
    const parsed = this.requireRange(range)
    return {
      range: this.formatRange(parsed),
      values: this.workpaper.getRangeValues(parsed),
      serialized: this.workpaper.getRangeSerialized(parsed),
      displayValues: this.readRangeDisplayValues(parsed),
    }
  }

  range(range: string): A1CellValue[][] {
    return this.readRange(range).values
  }

  setCell(address: string, value: A1CellInput): WorkPaperChange[] {
    const parsed = this.requireCellAddress(address)
    this.assertWritable(parsed)
    this.assertFormulaCanParse(value, parsed)
    return this.workpaper.setCellContents(parsed, value)
  }

  set(address: string, value: A1CellInput): WorkPaperChange[] {
    return this.setCell(address, value)
  }

  readMany(addresses: readonly string[]): Readonly<Record<string, A1CellRead>> {
    const reads: Record<string, A1CellRead> = {}
    for (const address of addresses) {
      const read = this.readCell(address)
      reads[read.address] = read
    }
    return reads
  }

  setCells(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[] {
    const changes: WorkPaperChange[] = []
    for (const [address, value] of Object.entries(edits)) {
      changes.push(...this.setCell(address, value))
    }
    return changes
  }

  setMany(edits: Readonly<Record<string, A1CellInput>>): WorkPaperChange[] {
    return this.setCells(edits)
  }

  setCellAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof {
    const requireReadbackChange = options.requireReadbackChange ?? true
    const includeConfig = options.includeConfig ?? true
    const before = this.readCell(address)
    const beforeReadback = this.readRange(options.readbackRange)
    const changes = this.setCell(address, value)
    const after = this.readCell(address)
    const afterReadback = this.readRange(options.readbackRange)
    const serializedDocument = this.serialize({ includeConfig })
    const restored = restoreA1WorkPaper(serializedDocument, this.facadeOptions())

    try {
      const restoredReadback = restored.readRange(options.readbackRange)
      const persistedDocumentBytes = new TextEncoder().encode(serializedDocument).byteLength
      const readbackChanged = !sameJson(rangeComparison(beforeReadback), rangeComparison(afterReadback))
      const restoredReadbackMatchesAfter = sameJson(rangeComparison(afterReadback), rangeComparison(restoredReadback))

      return {
        editedCell: after.address,
        readbackRange: afterReadback.range,
        before,
        after,
        beforeReadback,
        afterReadback,
        restoredReadback,
        persistedDocumentBytes,
        checks: {
          readbackChanged,
          restoredReadbackMatchesAfter,
          previousSerialized: before.serialized,
          newSerialized: after.serialized,
        },
        verified: persistedDocumentBytes > 0 && restoredReadbackMatchesAfter && (!requireReadbackChange || readbackChanged),
        limitations: [...this.limitations, ...(options.limitations ?? [])],
        changes,
        ...(options.includeSerializedDocument ? { serializedDocument } : {}),
      }
    } finally {
      restored.dispose()
    }
  }

  editAndReadback(address: string, value: A1CellInput, options: A1SetCellAndReadbackOptions): A1SetCellAndReadbackProof {
    return this.setCellAndReadback(address, value, options)
  }

  exportDocument(options: { readonly includeConfig?: boolean } = {}): PersistedWorkPaperDocument {
    return exportWorkPaperDocument(this.workpaper, options)
  }

  serialize(options: { readonly includeConfig?: boolean } = {}): string {
    return serializeWorkPaperDocument(this.exportDocument(options))
  }

  saveJson(options: { readonly includeConfig?: boolean } = {}): string {
    return this.serialize(options)
  }

  restoreJson(input: string | PersistedWorkPaperDocument): A1WorkPaper {
    return restoreA1WorkPaper(input, this.facadeOptions())
  }

  dispose(): void {
    this.workpaper.dispose()
  }

  private facadeOptions(): A1WorkPaperOptions {
    return {
      ...(this.defaultSheetName !== undefined ? { defaultSheetName: this.defaultSheetName } : {}),
      ...(this.writableSheets !== undefined ? { writableSheets: this.writableSheets } : {}),
      limitations: this.limitations,
    }
  }

  private readParsedCell(address: WorkPaperCellAddress): A1CellRead {
    const formula = this.workpaper.getCellFormula(address)
    return {
      address: this.formatCell(address),
      value: this.workpaper.getCellValue(address),
      serialized: this.workpaper.getCellSerialized(address),
      formula: formula ?? null,
      displayValue: this.workpaper.getCellDisplayValue(address),
    }
  }

  private readRangeDisplayValues(range: WorkPaperCellRange): string[][] {
    const displayValues: string[][] = []
    for (let row = range.start.row; row <= range.end.row; row += 1) {
      const rowValues: string[] = []
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        rowValues.push(
          this.workpaper.getCellDisplayValue({
            sheet: range.start.sheet,
            row,
            col,
          }),
        )
      }
      displayValues.push(rowValues)
    }
    return displayValues
  }

  private requireCellAddress(address: string): WorkPaperCellAddress {
    this.assertDefaultSheetAllowsUnqualifiedAddress(address, 'cell address')
    const parsed = this.workpaper.simpleCellAddressFromString(address, this.defaultSheetId())
    if (parsed === undefined) {
      throw new Error(`Invalid WorkPaper cell address: ${address}`)
    }
    return parsed
  }

  private requireRange(range: string): WorkPaperCellRange {
    this.assertDefaultSheetAllowsUnqualifiedAddress(range, 'range')
    const parsedRange = this.workpaper.simpleCellRangeFromString(range, this.defaultSheetId())
    if (parsedRange !== undefined) {
      return parsedRange
    }
    const parsedCell = this.workpaper.simpleCellAddressFromString(range, this.defaultSheetId())
    if (parsedCell !== undefined) {
      return {
        start: parsedCell,
        end: parsedCell,
      }
    }
    throw new Error(`Invalid WorkPaper range: ${range}`)
  }

  private assertDefaultSheetAllowsUnqualifiedAddress(value: string, kind: string): void {
    if (this.defaultSheetName === undefined && !value.includes('!')) {
      throw new Error(`Sheet-qualified WorkPaper ${kind} required: ${value}`)
    }
  }

  private defaultSheetId(): number | undefined {
    if (this.defaultSheetName === undefined) {
      return undefined
    }
    const sheetId = this.workpaper.getSheetId(this.defaultSheetName)
    if (sheetId === undefined) {
      throw new Error(`Expected default sheet "${this.defaultSheetName}" to exist`)
    }
    return sheetId
  }

  private assertWritable(address: WorkPaperCellAddress): void {
    if (this.writableSheets === undefined) {
      return
    }
    const sheetName = this.workpaper.getSheetName(address.sheet)
    if (sheetName === undefined || !this.writableSheets.includes(sheetName)) {
      throw new Error(`Sheet "${sheetName ?? address.sheet}" is not writable`)
    }
  }

  private assertFormulaCanParse(value: A1CellInput, address: WorkPaperCellAddress): void {
    if (typeof value !== 'string' || !value.startsWith('=')) {
      return
    }
    if (!this.workpaper.validateFormula(value)) {
      throw new Error(`Invalid WorkPaper formula for ${this.formatCell(address)}: ${value}`)
    }
  }

  private formatCell(address: WorkPaperCellAddress): string {
    return this.workpaper.simpleCellAddressToString(address, {
      includeSheetName: true,
    })
  }

  private formatRange(range: WorkPaperCellRange): string {
    if (sameAddress(range.start, range.end)) {
      return this.formatCell(range.start)
    }
    return this.workpaper.simpleCellRangeToString(range, {
      includeSheetName: true,
    })
  }
}

function rangeComparison(read: A1RangeRead): unknown {
  return {
    displayValues: read.displayValues,
    serialized: read.serialized,
  }
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function sameAddress(left: WorkPaperCellAddress, right: WorkPaperCellAddress): boolean {
  return left.sheet === right.sheet && left.row === right.row && left.col === right.col
}
