import type { LiteralInput } from '@bilig/protocol'
import { filledUint32Array, growFloat64Array, growInt16Array, growInt32Array, growUint32Array } from './xlsx-large-simple-array-storage.js'
import {
  binarySearchUint32Prefix,
  canStoreInt8Number,
  canStoreInt16Number,
  canStoreInt32Number,
} from './xlsx-large-simple-arena-helpers.js'
import {
  initialSparseIntegerCapacity,
  noPoolId,
  valueKindBoolean,
  valueKindEmpty,
  valueKindEmptyString,
  valueKindInteger,
  valueKindNull,
  valueKindNumber,
  valueKindSharedStringRef,
  valueKindSmallInteger,
  valueKindString,
  valueKindTinyInteger,
} from './xlsx-large-simple-arena-constants.js'
import {
  canStoreStringIdInStorage,
  createStringIdStorage,
  growStringIdStorage,
  ImportedWorkbookArenaCoordinateBase,
  maxNarrowStringId,
  normalizeStoredStringId,
  stringNoPoolIdForStorage,
  type StringIdStorage,
  widenStringIdStorage,
} from './xlsx-large-simple-arena-coordinate-base.js'
import type { ImportedWorkbookArenaDedupeMode } from './xlsx-large-simple-arena-types.js'

const sparseStringDenseThresholdFloor = 1_000_000

export abstract class ImportedWorkbookArenaBase extends ImportedWorkbookArenaCoordinateBase {
  protected addValue(index: number, value: LiteralInput | undefined): void {
    if (value === undefined) {
      this.valueKinds[index] = valueKindEmpty
      return
    }
    if (value === null) {
      this.valueKinds[index] = valueKindNull
      return
    }
    if (typeof value === 'number') {
      if (canStoreInt8Number(value)) {
        this.valueKinds[index] = valueKindTinyInteger
        this.ensureTinyIntegerValueStorage()[index] = value
        return
      }
      if (canStoreInt16Number(value)) {
        this.valueKinds[index] = valueKindSmallInteger
        this.storeSmallIntegerValue(index, value)
        return
      }
      if (canStoreInt32Number(value)) {
        this.valueKinds[index] = valueKindInteger
        this.storeIntegerValue(index, value)
        return
      }
      this.valueKinds[index] = valueKindNumber
      this.storeNumberValue(index, value)
      return
    }
    if (typeof value === 'boolean') {
      this.valueKinds[index] = valueKindBoolean
      this.ensureBooleanValueStorage()[index] = value ? 1 : 0
      return
    }
    if (value === '') {
      this.valueKinds[index] = valueKindEmptyString
      return
    }
    this.valueKinds[index] = valueKindString
    this.stringValueCount += 1
    this.storeStringId(index, this.internString(value))
  }

  protected materializeValue(index: number): LiteralInput | undefined {
    const valueKind = this.valueKinds[index] ?? valueKindEmpty
    switch (valueKind) {
      case valueKindNumber:
        return this.numberValueAt(index)
      case valueKindTinyInteger:
        return this.tinyIntegerValues?.[index]
      case valueKindSmallInteger:
        return this.smallIntegerValueAt(index)
      case valueKindInteger:
        return this.integerValueAt(index)
      case valueKindString: {
        const stringId = this.stringIdAt(index)
        return stringId === noPoolId ? undefined : this.strings[stringId]
      }
      case valueKindSharedStringRef: {
        const sharedStringIndex = this.sharedStringIndexAt(index)
        return sharedStringIndex === noPoolId ? undefined : this.sharedStrings?.[sharedStringIndex]?.text
      }
      case valueKindBoolean:
        return (this.booleanValues?.[index] ?? 0) === 1
      case valueKindNull:
        return null
      case valueKindEmptyString:
        return ''
      default:
        return undefined
    }
  }

  protected hasCellsForSheet(sheetIndex: number): boolean {
    return this.sheetIndexes !== undefined || this.sheetIndex === sheetIndex
  }

  protected cellBelongsToSheet(index: number, sheetIndex: number): boolean {
    return this.sheetIndexes ? this.sheetIndexes[index] === sheetIndex : this.sheetIndex === sheetIndex
  }

  protected ensureStringIdStorage(): StringIdStorage {
    if (this.stringIds) {
      return this.stringIds
    }
    if (this.sparseStringCellIndexes && this.sparseStringIds) {
      const output = createStringIdStorage(this.valueKinds.length, this.sparseStringIds instanceof Uint32Array)
      for (let index = 0; index < this.sparseStringCount; index += 1) {
        const cellIndex = this.sparseStringCellIndexes[index] ?? -1
        if (cellIndex >= 0 && cellIndex < output.length) {
          output[cellIndex] = this.sparseStringIds[index] ?? stringNoPoolIdForStorage(output)
        }
      }
      this.sparseStringCellIndexes = undefined
      this.sparseStringIds = undefined
      this.sparseStringCount = 0
      this.stringIds = output
      return this.stringIds
    }
    this.stringIds = createStringIdStorage(this.valueKinds.length)
    return this.stringIds
  }

  protected ensureNumberValueStorage(): Float64Array<ArrayBuffer> {
    if (this.numberValues) {
      return this.numberValues
    }
    this.numberValues = new Float64Array(this.valueKinds.length)
    this.numberValues.fill(Number.NaN)
    const sparseIndexes = this.sparseNumberCellIndexes
    const sparseValues = this.sparseNumberValues
    if (sparseIndexes && sparseValues) {
      for (let offset = 0; offset < this.sparseNumberCount; offset += 1) {
        const cellIndex = sparseIndexes[offset] ?? -1
        if (cellIndex >= 0 && cellIndex < this.numberValues.length) {
          this.numberValues[cellIndex] = sparseValues[offset] ?? Number.NaN
        }
      }
    }
    this.sparseNumberCellIndexes = undefined
    this.sparseNumberValues = undefined
    this.sparseNumberCount = 0
    this.moveSharedStringIndexesToNumberValues()
    return this.numberValues
  }

  protected storeNumberValue(index: number, value: number): void {
    if (this.numberValues) {
      this.numberValues[index] = value
      return
    }
    if (this.sparseNumberCount >= this.sparseNumberDenseThreshold()) {
      this.ensureNumberValueStorage()[index] = value
      return
    }
    if (!this.sparseNumberCellIndexes || !this.sparseNumberValues) {
      this.sparseNumberCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseNumberValues = new Float64Array(initialSparseIntegerCapacity)
      this.sparseNumberValues.fill(Number.NaN)
    } else if (this.sparseNumberCount >= this.sparseNumberCellIndexes.length) {
      const nextCapacity = this.sparseNumberCellIndexes.length * 2
      this.sparseNumberCellIndexes = growUint32Array(this.sparseNumberCellIndexes, nextCapacity)
      this.sparseNumberValues = growFloat64Array(this.sparseNumberValues, nextCapacity)
    }
    this.sparseNumberCellIndexes[this.sparseNumberCount] = index
    this.sparseNumberValues[this.sparseNumberCount] = value
    this.sparseNumberCount += 1
  }

  protected numberValueAt(index: number): number | undefined {
    if (this.numberValues) {
      return this.numberValues[index]
    }
    const sparseIndexes = this.sparseNumberCellIndexes
    const sparseValues = this.sparseNumberValues
    if (!sparseIndexes || !sparseValues || this.sparseNumberCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseNumberCount, index)
    return offset === -1 ? undefined : sparseValues[offset]
  }

  protected ensureTinyIntegerValueStorage(): Int8Array<ArrayBuffer> {
    if (this.tinyIntegerValues) {
      return this.tinyIntegerValues
    }
    this.tinyIntegerValues = new Int8Array(this.valueKinds.length)
    return this.tinyIntegerValues
  }

  protected storeSmallIntegerValue(index: number, value: number): void {
    if (this.smallIntegerValues) {
      this.smallIntegerValues[index] = value
      return
    }
    if (this.sparseSmallIntegerCount >= this.sparseIntegerDenseThreshold()) {
      this.ensureSmallIntegerValueStorage()[index] = value
      return
    }
    if (!this.sparseSmallIntegerCellIndexes || !this.sparseSmallIntegerValues) {
      this.sparseSmallIntegerCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseSmallIntegerValues = new Int16Array(initialSparseIntegerCapacity)
    } else if (this.sparseSmallIntegerCount >= this.sparseSmallIntegerCellIndexes.length) {
      const nextCapacity = this.sparseSmallIntegerCellIndexes.length * 2
      this.sparseSmallIntegerCellIndexes = growUint32Array(this.sparseSmallIntegerCellIndexes, nextCapacity)
      this.sparseSmallIntegerValues = growInt16Array(this.sparseSmallIntegerValues, nextCapacity)
    }
    this.sparseSmallIntegerCellIndexes[this.sparseSmallIntegerCount] = index
    this.sparseSmallIntegerValues[this.sparseSmallIntegerCount] = value
    this.sparseSmallIntegerCount += 1
  }

  protected ensureSmallIntegerValueStorage(): Int16Array<ArrayBuffer> {
    if (this.smallIntegerValues) {
      return this.smallIntegerValues
    }
    this.smallIntegerValues = new Int16Array(this.valueKinds.length)
    const sparseIndexes = this.sparseSmallIntegerCellIndexes
    const sparseValues = this.sparseSmallIntegerValues
    if (sparseIndexes && sparseValues) {
      for (let offset = 0; offset < this.sparseSmallIntegerCount; offset += 1) {
        const cellIndex = sparseIndexes[offset] ?? -1
        if (cellIndex >= 0 && cellIndex < this.smallIntegerValues.length) {
          this.smallIntegerValues[cellIndex] = sparseValues[offset] ?? 0
        }
      }
    }
    this.sparseSmallIntegerCellIndexes = undefined
    this.sparseSmallIntegerValues = undefined
    this.sparseSmallIntegerCount = 0
    return this.smallIntegerValues
  }

  protected smallIntegerValueAt(index: number): number | undefined {
    if (this.smallIntegerValues) {
      return this.smallIntegerValues[index]
    }
    const sparseIndexes = this.sparseSmallIntegerCellIndexes
    const sparseValues = this.sparseSmallIntegerValues
    if (!sparseIndexes || !sparseValues || this.sparseSmallIntegerCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseSmallIntegerCount, index)
    return offset === -1 ? undefined : sparseValues[offset]
  }

  protected storeIntegerValue(index: number, value: number): void {
    if (this.integerValues) {
      this.integerValues[index] = value
      return
    }
    if (this.sparseIntegerCount >= this.sparseIntegerDenseThreshold()) {
      this.ensureIntegerValueStorage()[index] = value
      return
    }
    if (!this.sparseIntegerCellIndexes || !this.sparseIntegerValues) {
      this.sparseIntegerCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseIntegerValues = new Int32Array(initialSparseIntegerCapacity)
    } else if (this.sparseIntegerCount >= this.sparseIntegerCellIndexes.length) {
      const nextCapacity = this.sparseIntegerCellIndexes.length * 2
      this.sparseIntegerCellIndexes = growUint32Array(this.sparseIntegerCellIndexes, nextCapacity)
      this.sparseIntegerValues = growInt32Array(this.sparseIntegerValues, nextCapacity)
    }
    this.sparseIntegerCellIndexes[this.sparseIntegerCount] = index
    this.sparseIntegerValues[this.sparseIntegerCount] = value
    this.sparseIntegerCount += 1
  }

  protected ensureIntegerValueStorage(): Int32Array<ArrayBuffer> {
    if (this.integerValues) {
      return this.integerValues
    }
    this.integerValues = new Int32Array(this.valueKinds.length)
    const sparseIndexes = this.sparseIntegerCellIndexes
    const sparseValues = this.sparseIntegerValues
    if (sparseIndexes && sparseValues) {
      for (let offset = 0; offset < this.sparseIntegerCount; offset += 1) {
        const cellIndex = sparseIndexes[offset] ?? -1
        if (cellIndex >= 0 && cellIndex < this.integerValues.length) {
          this.integerValues[cellIndex] = sparseValues[offset] ?? 0
        }
      }
    }
    this.sparseIntegerCellIndexes = undefined
    this.sparseIntegerValues = undefined
    this.sparseIntegerCount = 0
    return this.integerValues
  }

  protected integerValueAt(index: number): number | undefined {
    if (this.integerValues) {
      return this.integerValues[index]
    }
    const sparseIndexes = this.sparseIntegerCellIndexes
    const sparseValues = this.sparseIntegerValues
    if (!sparseIndexes || !sparseValues || this.sparseIntegerCount === 0) {
      return undefined
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseIntegerCount, index)
    return offset === -1 ? undefined : sparseValues[offset]
  }

  protected sparseIntegerDenseThreshold(): number {
    return Math.max(initialSparseIntegerCapacity, this.valueKinds.length >>> 2)
  }

  protected sparseNumberDenseThreshold(): number {
    return Math.max(initialSparseIntegerCapacity, Math.floor((this.valueKinds.length * 2) / 3))
  }

  protected sparseStringDenseThreshold(): number {
    return Math.max(sparseStringDenseThresholdFloor, this.valueKinds.length >>> 2)
  }

  protected ensureFormulaIdStorage(): Uint32Array<ArrayBuffer> {
    if (this.formulaIds) {
      return this.formulaIds
    }
    this.formulaIds = filledUint32Array(this.valueKinds.length, noPoolId)
    return this.formulaIds
  }

  protected ensureBooleanValueStorage(): Uint8Array<ArrayBuffer> {
    if (this.booleanValues) {
      return this.booleanValues
    }
    this.booleanValues = new Uint8Array(this.valueKinds.length)
    return this.booleanValues
  }

  protected storeSharedStringIndex(index: number, sharedStringIndex: number): void {
    this.sharedStringRefCount += 1
    if (this.numberValues) {
      this.sharedStringRefsInNumberValues = true
      this.numberValues[index] = sharedStringIndex
      return
    }
    this.storeStringId(index, sharedStringIndex)
  }

  protected sharedStringIndexAt(index: number): number {
    if (this.sharedStringRefsInNumberValues) {
      const value = this.numberValues?.[index]
      if (value !== undefined && !Number.isNaN(value)) {
        return Math.trunc(value)
      }
    }
    return this.stringIdAt(index)
  }

  protected moveSharedStringIndexesToNumberValues(): void {
    if (!this.numberValues || this.sharedStringRefCount === 0 || this.stringValueCount > 0) {
      return
    }
    if (this.stringIds) {
      for (let index = 0; index < this.length; index += 1) {
        if ((this.valueKinds[index] ?? valueKindEmpty) === valueKindSharedStringRef) {
          this.numberValues[index] = normalizeStoredStringId(this.stringIds[index], this.stringIds)
        }
      }
      this.stringIds = undefined
      this.sharedStringRefsInNumberValues = true
      return
    }
    const sparseIndexes = this.sparseStringCellIndexes
    const sparseIds = this.sparseStringIds
    if (!sparseIndexes || !sparseIds || this.sparseStringCount === 0) {
      return
    }
    for (let offset = 0; offset < this.sparseStringCount; offset += 1) {
      const index = sparseIndexes[offset] ?? -1
      if (index >= 0 && index < this.length && (this.valueKinds[index] ?? valueKindEmpty) === valueKindSharedStringRef) {
        this.numberValues[index] = normalizeStoredStringId(sparseIds[offset], sparseIds)
      }
    }
    this.sparseStringCellIndexes = undefined
    this.sparseStringIds = undefined
    this.sparseStringCount = 0
    this.sharedStringRefsInNumberValues = true
  }

  protected storeStringId(index: number, stringId: number): void {
    if (this.stringIds) {
      if (!canStoreStringIdInStorage(this.stringIds, stringId)) {
        this.stringIds = widenStringIdStorage(this.stringIds)
      }
      this.stringIds[index] = stringId
      return
    }
    if (this.sparseStringCellIndexes && this.sparseStringIds) {
      const existingOffset = binarySearchUint32Prefix(this.sparseStringCellIndexes, this.sparseStringCount, index)
      if (existingOffset !== -1) {
        if (!canStoreStringIdInStorage(this.sparseStringIds, stringId)) {
          this.sparseStringIds = widenStringIdStorage(this.sparseStringIds)
        }
        this.sparseStringIds[existingOffset] = stringId
        return
      }
    }
    if (this.sparseStringCount >= this.sparseStringDenseThreshold()) {
      let stringIds = this.ensureStringIdStorage()
      if (!canStoreStringIdInStorage(stringIds, stringId)) {
        stringIds = widenStringIdStorage(stringIds)
        this.stringIds = stringIds
      }
      stringIds[index] = stringId
      return
    }
    if (!this.sparseStringCellIndexes || !this.sparseStringIds) {
      this.sparseStringCellIndexes = new Uint32Array(initialSparseIntegerCapacity)
      this.sparseStringIds = createStringIdStorage(initialSparseIntegerCapacity, stringId > maxNarrowStringId)
    } else {
      if (!canStoreStringIdInStorage(this.sparseStringIds, stringId)) {
        this.sparseStringIds = widenStringIdStorage(this.sparseStringIds)
      }
      if (this.sparseStringCount >= this.sparseStringCellIndexes.length) {
        const nextCapacity = this.sparseStringCellIndexes.length * 2
        this.sparseStringCellIndexes = growUint32Array(this.sparseStringCellIndexes, nextCapacity)
        this.sparseStringIds = growStringIdStorage(this.sparseStringIds, nextCapacity)
      }
    }
    const insertOffset = this.sparseStringInsertOffset(index)
    if (insertOffset < this.sparseStringCount) {
      this.sparseStringCellIndexes.copyWithin(insertOffset + 1, insertOffset, this.sparseStringCount)
      this.sparseStringIds.copyWithin(insertOffset + 1, insertOffset, this.sparseStringCount)
    }
    this.sparseStringCellIndexes[insertOffset] = index
    this.sparseStringIds[insertOffset] = stringId
    this.sparseStringCount += 1
  }

  protected sparseStringInsertOffset(index: number): number {
    const sparseIndexes = this.sparseStringCellIndexes
    if (!sparseIndexes || this.sparseStringCount === 0) {
      return 0
    }
    let low = 0
    let high = this.sparseStringCount
    while (low < high) {
      const mid = (low + high) >>> 1
      if ((sparseIndexes[mid] ?? 0) < index) {
        low = mid + 1
      } else {
        high = mid
      }
    }
    return low
  }

  protected stringIdAt(index: number): number {
    if (this.stringIds) {
      return normalizeStoredStringId(this.stringIds[index], this.stringIds)
    }
    const sparseIndexes = this.sparseStringCellIndexes
    const sparseIds = this.sparseStringIds
    if (!sparseIndexes || !sparseIds || this.sparseStringCount === 0) {
      return noPoolId
    }
    const offset = binarySearchUint32Prefix(sparseIndexes, this.sparseStringCount, index)
    return offset === -1 ? noPoolId : normalizeStoredStringId(sparseIds[offset], sparseIds)
  }

  protected internString(value: string): number {
    const interned = this.internValue(value, this.stringDedupeMode)
    if (interned === null) {
      const next = this.strings.length
      this.strings.push(value)
      return next
    }
    const existing = this.stringIdsByValue.get(interned)
    if (existing !== undefined) {
      return existing
    }
    const next = this.strings.length
    this.strings.push(interned)
    this.stringIdsByValue.set(interned, next)
    if (this.stringDedupeMode === 'bounded') {
      this.rememberBoundedDedupeKey(this.stringIdsByValue, this.stringDedupeKeys, 'string', interned)
    }
    return next
  }

  protected internFormula(value: string): number {
    const interned = this.internValue(value, this.formulaDedupeMode)
    if (interned === null) {
      const next = this.formulas.length
      this.formulas.push(value)
      return next
    }
    const existing = this.formulaIdsByValue.get(interned)
    if (existing !== undefined) {
      return existing
    }
    const next = this.formulas.length
    this.formulas.push(interned)
    this.formulaIdsByValue.set(interned, next)
    if (this.formulaDedupeMode === 'bounded') {
      this.rememberBoundedDedupeKey(this.formulaIdsByValue, this.formulaDedupeKeys, 'formula', interned)
    }
    return next
  }

  protected internValue(value: string, mode: ImportedWorkbookArenaDedupeMode): string | null {
    if (mode === false) {
      return null
    }
    if (mode === 'bounded') {
      return this.stringPool?.internBounded(value, this.dedupeMaxEntries) ?? value
    }
    return this.stringPool?.intern(value) ?? value
  }

  protected rememberBoundedDedupeKey(map: Map<string, number>, keys: string[], kind: 'string' | 'formula', key: string): void {
    keys.push(key)
    let evictionIndex = kind === 'string' ? this.stringDedupeEvictionIndex : this.formulaDedupeEvictionIndex
    while (keys.length - evictionIndex > this.dedupeMaxEntries) {
      const evicted = keys[evictionIndex]
      evictionIndex += 1
      if (evicted !== undefined) {
        map.delete(evicted)
      }
    }
    if (evictionIndex > this.dedupeMaxEntries && evictionIndex * 2 > keys.length) {
      keys.splice(0, evictionIndex)
      evictionIndex = 0
    }
    if (kind === 'string') {
      this.stringDedupeEvictionIndex = evictionIndex
    } else {
      this.formulaDedupeEvictionIndex = evictionIndex
    }
  }

  protected compactSparseStringIds(): void {
    if (this.sparseStringCellIndexes && this.sparseStringIds) {
      if (this.sparseStringCount === 0) {
        this.sparseStringCellIndexes = undefined
        this.sparseStringIds = undefined
        return
      }
      if (this.sparseStringCellIndexes.length !== this.sparseStringCount) {
        this.sparseStringCellIndexes = this.sparseStringCellIndexes.slice(0, this.sparseStringCount)
      }
      if (this.sparseStringIds.length !== this.sparseStringCount) {
        this.sparseStringIds = this.sparseStringIds.slice(0, this.sparseStringCount)
      }
      return
    }
    const denseStringIds = this.stringIds
    if (!denseStringIds || this.length === 0) {
      return
    }
    let retainedCount = 0
    for (let index = 0; index < this.length; index += 1) {
      if (normalizeStoredStringId(denseStringIds[index], denseStringIds) !== noPoolId) {
        retainedCount += 1
      }
    }
    if (retainedCount === 0) {
      this.stringIds = undefined
      this.sparseStringCellIndexes = undefined
      this.sparseStringIds = undefined
      this.sparseStringCount = 0
      return
    }
    if (retainedCount * 2 >= this.length) {
      return
    }
    const indexes = new Uint32Array(retainedCount)
    const ids = createStringIdStorage(retainedCount, denseStringIds instanceof Uint32Array)
    let outputIndex = 0
    for (let index = 0; index < this.length; index += 1) {
      const stringId = normalizeStoredStringId(denseStringIds[index], denseStringIds)
      if (stringId === noPoolId) {
        continue
      }
      indexes[outputIndex] = index
      ids[outputIndex] = stringId
      outputIndex += 1
    }
    this.stringIds = undefined
    this.sparseStringCellIndexes = indexes
    this.sparseStringIds = ids
    this.sparseStringCount = retainedCount
  }

  protected snapshotStringIds(): StringIdStorage | undefined {
    if (this.stringIds) {
      return this.stringIds.subarray(0, this.length)
    }
    const sparseIndexes = this.sparseStringCellIndexes
    const sparseIds = this.sparseStringIds
    if (!sparseIndexes || !sparseIds || this.sparseStringCount === 0) {
      return undefined
    }
    const output = createStringIdStorage(this.length, sparseIds instanceof Uint32Array)
    for (let index = 0; index < this.sparseStringCount; index += 1) {
      const cellIndex = sparseIndexes[index] ?? -1
      if (cellIndex >= 0 && cellIndex < output.length) {
        output[cellIndex] = sparseIds[index] ?? stringNoPoolIdForStorage(output)
      }
    }
    return output
  }
}
