import { translateFormulaReferences } from '@bilig/formula'
import { normalizeImportedFormulaSource } from './xlsx-formula-translation.js'
import { formulaReferencesExternalWorkbook, formulaReferencesVolatileFunction } from './xlsx-import-warnings.js'
import type { ImportedWorkbookArena } from './xlsx-large-simple-arena.js'

const initialFormulaCapacity = 128
const boundedRawFormulaDedupeMaxEntries = 8192
const noPoolId = 0xffffffff
const formulaTypeNormal = 0
const formulaTypeShared = 1

export interface LargeSimpleFormulaNumericRecords {
  readonly cellIndexes: Uint32Array
  readonly rows: Uint32Array
  readonly columns: Uint16Array | Uint32Array
  readonly typeCodes: Uint8Array
  readonly sharedIndexes: Uint32Array
}

export class LargeSimpleFormulaRecords {
  private cellIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialFormulaCapacity)
  private rows: Uint32Array<ArrayBuffer> | undefined
  private columns: Uint32Array<ArrayBuffer> | undefined
  private typeCodes: Uint8Array<ArrayBuffer> | undefined
  private sharedIndexes: Uint32Array<ArrayBuffer> | undefined
  private rawFormulaIds: Uint32Array<ArrayBuffer> = filledUint32Array(initialFormulaCapacity, noPoolId)
  private readonly rawFormulas: string[] = []
  private readonly rawFormulaIdsByValue = new Map<string, number>()
  private readonly boundedRawFormulaKeys: string[] = []
  private boundedRawFormulaEvictionIndex = 0
  private readonly normalizedFormulas: (string | null | undefined)[] = []
  private length = 0

  constructor(private readonly allowUnsupportedFormulaText = false) {}

  get count(): number {
    return this.length
  }

  get rawFormulaPoolCount(): number {
    return this.rawFormulas.length
  }

  retainedStorageByteLength(): number {
    return (
      this.cellIndexes.byteLength +
      (this.rows?.byteLength ?? 0) +
      (this.columns?.byteLength ?? 0) +
      (this.typeCodes?.byteLength ?? 0) +
      (this.sharedIndexes?.byteLength ?? 0) +
      this.rawFormulaIds.byteLength
    )
  }

  add(cellIndex: number, row: number, column: number, typeCode: number, sharedIndex: number | null, rawFormulaText: string): void {
    this.ensureCapacity(this.length + 1)
    const index = this.length
    this.length += 1
    this.cellIndexes[index] = cellIndex
    this.rawFormulaIds[index] = this.internRawFormula(rawFormulaText)
    if (typeCode === formulaTypeShared) {
      this.ensureSharedFormulaStorage()
      this.rows![index] = row
      this.columns![index] = column
      this.typeCodes![index] = typeCode
      this.sharedIndexes![index] = sharedIndex ?? noPoolId
    }
  }

  addRawFormula(rawFormulaText: string): void {
    this.ensureCapacity(this.length + 1)
    const index = this.length
    this.length += 1
    this.rawFormulaIds[index] = this.internRawFormula(rawFormulaText)
  }

  hydrateNumericRecords(records: LargeSimpleFormulaNumericRecords): void {
    const count = records.cellIndexes.length
    if (count === 0) {
      return
    }
    this.ensureCapacity(count)
    this.cellIndexes.set(records.cellIndexes, 0)
    if (containsSharedFormulaRecords(records.typeCodes)) {
      this.ensureSharedFormulaStorage()
      this.rows!.set(records.rows, 0)
      this.columns!.set(records.columns, 0)
      this.typeCodes!.set(records.typeCodes, 0)
      this.sharedIndexes!.set(records.sharedIndexes, 0)
    }
  }

  resolveIntoArena(arena: ImportedWorkbookArena, numericRecords?: LargeSimpleFormulaNumericRecords): boolean {
    this.releaseDedupeScratch()
    const length = numericRecords?.cellIndexes.length ?? this.length
    const sharedBases = new Map<number, SharedFormulaBase>()
    for (let index = 0; index < length; index += 1) {
      if (!this.isSharedFormula(index, numericRecords) || !this.hasRawFormulaText(index)) {
        continue
      }
      const normalized = this.normalizedFormulaAt(index)
      const sharedIndex = this.sharedIndexAt(index, numericRecords)
      if (normalized === null || sharedIndex === noPoolId) {
        return false
      }
      sharedBases.set(sharedIndex, {
        row: this.rowAt(index, numericRecords),
        column: this.columnAt(index, numericRecords),
        formula: normalized,
      })
      arena.setFormula(this.cellIndexAt(index, numericRecords), normalized)
    }

    for (let index = 0; index < length; index += 1) {
      if (this.isSharedFormula(index, numericRecords)) {
        if (this.hasRawFormulaText(index)) {
          continue
        }
        const sharedIndex = this.sharedIndexAt(index, numericRecords)
        const base = sharedIndex === noPoolId ? undefined : sharedBases.get(sharedIndex)
        if (!base) {
          return false
        }
        try {
          arena.setFormula(
            this.cellIndexAt(index, numericRecords),
            translateFormulaReferences(
              base.formula,
              this.rowAt(index, numericRecords) - base.row,
              this.columnAt(index, numericRecords) - base.column,
            ),
          )
        } catch {
          return false
        }
        continue
      }
      const normalized = this.normalizedFormulaAt(index)
      if (normalized === null) {
        return false
      }
      arena.setFormula(this.cellIndexAt(index, numericRecords), normalized)
    }
    return true
  }

  release(): void {
    this.cellIndexes = new Uint32Array(0)
    this.rows = undefined
    this.columns = undefined
    this.typeCodes = undefined
    this.sharedIndexes = undefined
    this.rawFormulaIds = new Uint32Array(0)
    this.rawFormulas.length = 0
    this.normalizedFormulas.length = 0
    this.length = 0
    this.releaseDedupeScratch()
  }

  private releaseDedupeScratch(): void {
    this.rawFormulaIdsByValue.clear()
    this.boundedRawFormulaKeys.length = 0
    this.boundedRawFormulaEvictionIndex = 0
  }

  private cellIndexAt(index: number, numericRecords: LargeSimpleFormulaNumericRecords | undefined): number {
    return numericRecords?.cellIndexes[index] ?? this.cellIndexes[index] ?? 0
  }

  private rowAt(index: number, numericRecords: LargeSimpleFormulaNumericRecords | undefined): number {
    return numericRecords?.rows[index] ?? this.rows?.[index] ?? 0
  }

  private columnAt(index: number, numericRecords: LargeSimpleFormulaNumericRecords | undefined): number {
    return numericRecords?.columns[index] ?? this.columns?.[index] ?? 0
  }

  private sharedIndexAt(index: number, numericRecords: LargeSimpleFormulaNumericRecords | undefined): number {
    return numericRecords?.sharedIndexes[index] ?? this.sharedIndexes?.[index] ?? noPoolId
  }

  private isSharedFormula(index: number, numericRecords?: LargeSimpleFormulaNumericRecords): boolean {
    return (numericRecords?.typeCodes[index] ?? this.typeCodes?.[index] ?? formulaTypeNormal) === formulaTypeShared
  }

  private hasRawFormulaText(index: number): boolean {
    return this.rawFormulaText(index).trim().length > 0
  }

  private rawFormulaText(index: number): string {
    const rawFormulaId = this.rawFormulaIds[index] ?? noPoolId
    return rawFormulaId === noPoolId ? '' : (this.rawFormulas[rawFormulaId] ?? '')
  }

  private normalizedFormulaAt(index: number): string | null {
    const rawFormulaId = this.rawFormulaIds[index] ?? noPoolId
    if (rawFormulaId === noPoolId) {
      return null
    }
    if (rawFormulaId < this.normalizedFormulas.length && this.normalizedFormulas[rawFormulaId] !== undefined) {
      return this.normalizedFormulas[rawFormulaId] ?? null
    }
    const normalized = normalizeLargeSimpleFormula(this.rawFormulas[rawFormulaId], this.allowUnsupportedFormulaText)
    this.normalizedFormulas[rawFormulaId] = normalized
    return normalized
  }

  private internRawFormula(value: string): number {
    const existing = this.rawFormulaIdsByValue.get(value)
    if (existing !== undefined) {
      return existing
    }
    const next = this.rawFormulas.length
    this.rawFormulas.push(value)
    this.rawFormulaIdsByValue.set(value, next)
    this.rememberBoundedRawFormulaKey(value)
    return next
  }

  private rememberBoundedRawFormulaKey(value: string): void {
    this.boundedRawFormulaKeys.push(value)
    while (this.boundedRawFormulaKeys.length - this.boundedRawFormulaEvictionIndex > boundedRawFormulaDedupeMaxEntries) {
      const evicted = this.boundedRawFormulaKeys[this.boundedRawFormulaEvictionIndex]
      this.boundedRawFormulaEvictionIndex += 1
      if (evicted !== undefined) {
        this.rawFormulaIdsByValue.delete(evicted)
      }
    }
    if (
      this.boundedRawFormulaEvictionIndex > boundedRawFormulaDedupeMaxEntries &&
      this.boundedRawFormulaEvictionIndex * 2 > this.boundedRawFormulaKeys.length
    ) {
      this.boundedRawFormulaKeys.splice(0, this.boundedRawFormulaEvictionIndex)
      this.boundedRawFormulaEvictionIndex = 0
    }
  }

  private ensureCapacity(nextLength: number): void {
    if (nextLength <= this.cellIndexes.length) {
      return
    }
    let nextCapacity = this.cellIndexes.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.cellIndexes = growUint32Array(this.cellIndexes, nextCapacity)
    if (this.rows) {
      this.rows = growUint32Array(this.rows, nextCapacity)
    }
    if (this.columns) {
      this.columns = growUint32Array(this.columns, nextCapacity)
    }
    if (this.typeCodes) {
      this.typeCodes = growUint8Array(this.typeCodes, nextCapacity)
    }
    if (this.sharedIndexes) {
      this.sharedIndexes = growUint32Array(this.sharedIndexes, nextCapacity, noPoolId)
    }
    this.rawFormulaIds = growUint32Array(this.rawFormulaIds, nextCapacity, noPoolId)
  }

  private ensureSharedFormulaStorage(): void {
    if (!this.rows) {
      this.rows = new Uint32Array(this.cellIndexes.length)
    }
    if (!this.columns) {
      this.columns = new Uint32Array(this.cellIndexes.length)
    }
    if (!this.typeCodes) {
      this.typeCodes = new Uint8Array(this.cellIndexes.length)
    }
    if (!this.sharedIndexes) {
      this.sharedIndexes = filledUint32Array(this.cellIndexes.length, noPoolId)
    }
  }
}

export function readLargeSimpleFormulaTypeCode(type: string | null): number {
  return type === 'shared' ? formulaTypeShared : formulaTypeNormal
}

export function parseLargeSimpleSharedFormulaIndex(value: string | null): number | null {
  if (!value || !/^[0-9]+$/u.test(value)) {
    return null
  }
  const index = Number(value)
  return Number.isSafeInteger(index) ? index : null
}

interface SharedFormulaBase {
  readonly row: number
  readonly column: number
  readonly formula: string
}

function containsSharedFormulaRecords(typeCodes: Uint8Array): boolean {
  for (const typeCode of typeCodes) {
    if (typeCode === formulaTypeShared) {
      return true
    }
  }
  return false
}

function normalizeLargeSimpleFormula(rawFormulaText: string | undefined, allowUnsupportedFormulaText: boolean): string | null {
  const decoded = rawFormulaText === undefined ? undefined : decodeXmlText(rawFormulaText).trim()
  if (decoded === undefined || decoded.length === 0) {
    return null
  }
  const formula = normalizeImportedFormulaSource(decoded)
  return !allowUnsupportedFormulaText &&
    (formulaReferencesExternalWorkbook(formula) || formulaReferencesVolatileFunction(formula) || formulaReferencesStructuredTable(formula))
    ? null
    : formula
}

function formulaReferencesStructuredTable(formula: string): boolean {
  return /\[[#@\w]/u.test(formula)
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function filledUint32Array(length: number, value: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(length)
  output.fill(value)
  return output
}

function growUint8Array(source: Uint8Array<ArrayBuffer>, nextCapacity: number): Uint8Array<ArrayBuffer> {
  const output = new Uint8Array(nextCapacity)
  output.set(source)
  return output
}

function growUint32Array(source: Uint32Array<ArrayBuffer>, nextCapacity: number, fillValue?: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(nextCapacity)
  output.set(source)
  if (fillValue !== undefined && nextCapacity > source.length) {
    output.fill(fillValue, source.length)
  }
  return output
}
