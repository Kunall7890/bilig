import {
  cellContentHasRichTextRun,
  doubleQuote,
  findClosingTag,
  findTagEnd,
  headlessTagCell,
  headlessTagDimension,
  headlessTagLocalName,
  headlessTagRow,
  headlessTagUnknown,
  attributeNameMatches,
  attributeValueMatches,
  decodePackedCellAddressBytes,
  isAsciiWhitespace,
  isSelfClosingTag,
  isXmlNameByte,
  lessThan,
  matchesDimensionTagName,
  metadataChildName,
  metadataCountMultiplier,
  packCellAddress,
  packedAddressColumnFactor,
  packedAddressColumn,
  packedAddressRow,
  readDimensionAddressRange,
  readNonNegativeIntegerFromRange,
  readPositiveIntegerAttributeFromTag,
  readXmlAttributeRangeFromTag,
  readXmlTagName,
  sheetMetadataKeysForHeadlessElement,
  skipAsciiWhitespace,
  slash,
  singleQuote,
} from './xlsx-large-simple-headless-worksheet-xml.js'
import { countLargeSimpleDataValidationsFromBytes } from './xlsx-large-simple-data-validation-byte-scan.js'
import { rowTagHasMetadataAttribute } from './xlsx-large-simple-row-metadata-scan.js'

const cellContentHasValue = 1 << 0
const cellContentHasFormula = 1 << 1
const emptyBytes = new Uint8Array(0)
const unsupportedWorksheetTagNames = new Set(['picture', 'sheetProtection'])
const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'controls',
  'dataValidations',
  'drawing',
  'headerFooter',
  'hyperlinks',
  'legacyDrawing',
  'mergeCells',
  'oleObjects',
  'pageMargins',
  'pageSetup',
  'printOptions',
  'rowBreaks',
  'sheetFormatPr',
  'tableParts',
])

export interface HeadlessLargeSimpleWorksheetScan {
  readonly sheetIndex: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount?: number
  readonly metadataKeys: readonly string[]
  readonly styleIndexes: ReadonlySet<number>
  readonly usesSharedStrings: boolean
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}

export interface HeadlessLargeSimpleWorksheetScanOptions {
  readonly hasSharedStrings: boolean
  readonly allowUnsupportedFeaturesForMetrics?: boolean
  readonly onRetainedBufferLength?: (length: number) => void
}

interface ActiveMetadataCount {
  readonly localName: string
  readonly childName: string | null
  readonly multiplier: number
  childCount: number
}

export function parseHeadlessLargeSimpleWorksheetFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: HeadlessLargeSimpleWorksheetScanOptions,
): HeadlessLargeSimpleWorksheetScan | null {
  const scanner = new HeadlessLargeSimpleWorksheetChunkScanner(sheetIndex, options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export async function parseHeadlessLargeSimpleWorksheetFromChunksAsync(
  readChunks: (onChunk: (chunk: Uint8Array) => boolean | void) => Promise<boolean>,
  sheetIndex: number,
  options: HeadlessLargeSimpleWorksheetScanOptions,
): Promise<HeadlessLargeSimpleWorksheetScan | null> {
  const scanner = new HeadlessLargeSimpleWorksheetChunkScanner(sheetIndex, options)
  if (!(await readChunks((chunk) => scanner.push(chunk)))) {
    return null
  }
  return scanner.finish()
}

class HeadlessLargeSimpleWorksheetChunkScanner {
  private buffer: Uint8Array = new Uint8Array()
  private index = 0
  private failed = false
  private rowCount = 0
  private columnCount = 0
  private cellCount = 0
  private valueCellCount = 0
  private formulaCellCount = 0
  private tableCount = 0
  private mergeCount = 0
  private conditionalFormatCount = 0
  private dataValidationCount = 0
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1
  private currentRow = -1
  private nextImplicitRow = 0
  private nextImplicitColumn = 0
  private cellContentFlags = 0
  private cellContentNextIndex = 0
  private cellPackedAddress = -1
  private cellStyleIndex: number | null = null
  private cellHasSharedStringType = false
  private cellHasInlineStringType = false
  private cellHasMetadataReference = false
  private usesSharedStrings = false
  private activeMetadata: ActiveMetadataCount | null = null
  private activeDataValidations = false
  private knownTagNameEnd = 0
  private readonly metadataKeys = new Set<string>()
  private readonly styleIndexes = new Set<number>()

  constructor(
    private readonly sheetIndex: number,
    private readonly options: HeadlessLargeSimpleWorksheetScanOptions,
  ) {}

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.append(chunk)
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
  }

  finish(): HeadlessLargeSimpleWorksheetScan | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    if (this.activeMetadata !== null) {
      this.failed = true
    }
    this.compact()
    this.reportRetainedBufferLength()
    return this.failed
      ? null
      : {
          sheetIndex: this.sheetIndex,
          cellCount: this.cellCount,
          valueCellCount: this.valueCellCount,
          formulaCellCount: this.formulaCellCount,
          tableCount: this.tableCount,
          mergeCount: this.mergeCount,
          conditionalFormatCount: this.conditionalFormatCount,
          dataValidationCount: this.dataValidationCount,
          metadataKeys: [...this.metadataKeys].toSorted(),
          styleIndexes: this.styleIndexes,
          usesSharedStrings: this.usesSharedStrings,
          rowCount: this.rowCount,
          columnCount: this.columnCount,
          usedRange:
            this.cellCount > 0
              ? {
                  startRow: this.minRow,
                  startColumn: this.minColumn,
                  endRow: this.maxRow,
                  endColumn: this.maxColumn,
                }
              : null,
        }
  }

  private append(chunk: Uint8Array): void {
    if (this.index === this.buffer.byteLength) {
      this.buffer = chunk
      this.index = 0
      return
    }
    const retained = this.buffer.subarray(this.index)
    const next = new Uint8Array(retained.byteLength + chunk.byteLength)
    next.set(retained)
    next.set(chunk, retained.byteLength)
    this.buffer = next
    this.index = 0
  }

  private compact(): void {
    if (this.index === 0) {
      return
    }
    if (this.index >= this.buffer.byteLength) {
      this.buffer = emptyBytes
      this.index = 0
      return
    }
    this.buffer = this.buffer.subarray(this.index)
    this.index = 0
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.byteLength) {
      if (this.activeMetadata !== null) {
        if (!this.processActiveMetadata(final)) {
          return
        }
        continue
      }
      if (this.activeDataValidations) {
        if (!this.processActiveDataValidations(final)) {
          return
        }
        continue
      }
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const knownTag = this.readKnownOpeningTagCode(this.index)
      let localName: string | null = null
      let nameEnd = this.knownTagNameEnd
      if (knownTag === headlessTagUnknown) {
        const tag = readXmlTagName(this.buffer, this.index + 1)
        if (!tag) {
          if (!final && this.index + 1 >= this.buffer.byteLength) {
            return
          }
          this.index += 1
          continue
        }
        localName = tag.localName
        nameEnd = tag.endIndex
      }
      if (knownTag !== headlessTagUnknown) {
        localName = headlessTagLocalName(knownTag)
      }
      if (localName === null) {
        if (!final && this.index + 1 >= this.buffer.byteLength) {
          return
        }
        this.index += 1
        continue
      }
      const tagEnd = findTagEnd(this.buffer, nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return
      }
      if (this.options.allowUnsupportedFeaturesForMetrics !== true && unsupportedWorksheetTagNames.has(localName)) {
        this.failed = true
        return
      }
      if (knownTag === headlessTagDimension || localName === 'dimension') {
        this.readDimension(nameEnd, tagEnd)
        this.index = tagEnd + 1
        continue
      }
      if (knownTag === headlessTagRow || localName === 'row') {
        this.readRow(nameEnd, tagEnd)
        this.index = tagEnd + 1
        continue
      }
      if (knownTag === headlessTagCell || localName === 'c') {
        if (!this.readCell(nameEnd, tagEnd, final)) {
          return
        }
        continue
      }
      if (metadataWorksheetTagNames.has(localName)) {
        if (localName === 'dataValidations') {
          if (!this.startDataValidations(tagEnd, final)) {
            return
          }
          continue
        }
        if (!this.countMetadataElement(localName, nameEnd, tagEnd, final)) {
          return
        }
        continue
      }
      this.index = tagEnd + 1
    }
  }

  private readDimension(nameEnd: number, tagEnd: number): void {
    const ref = readXmlAttributeRangeFromTag(this.buffer, nameEnd, tagEnd, 'ref')
    if (!ref) {
      return
    }
    const range = readDimensionAddressRange(this.buffer, ref.start, ref.end)
    if (!range) {
      return
    }
    this.rowCount = Math.max(this.rowCount, range.startRow + 1, range.endRow + 1)
    this.columnCount = Math.max(this.columnCount, range.startColumn + 1, range.endColumn + 1)
  }

  private readRow(nameEnd: number, tagEnd: number): void {
    const row = readPositiveIntegerAttributeFromTag(this.buffer, nameEnd, tagEnd, 'r')
    this.currentRow = row === null ? this.nextImplicitRow : row - 1
    this.nextImplicitRow = this.currentRow + 1
    this.nextImplicitColumn = 0
    if (rowTagHasMetadataAttribute(this.buffer, nameEnd, tagEnd)) {
      this.metadataKeys.add('rowMetadata')
      this.metadataKeys.add('rows')
    }
  }

  private readCell(nameEnd: number, tagEnd: number, final: boolean): boolean {
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    let contentFlags = 0
    let contentNextIndex = contentStart
    if (!selfClosing && !this.readCellContentSummary(contentStart)) {
      if (final) {
        this.failed = true
      }
      return false
    }
    if (!selfClosing) {
      contentFlags = this.cellContentFlags
      contentNextIndex = this.cellContentNextIndex
    }
    if (!this.readCellTagAttributes(nameEnd, tagEnd) || (!this.options.hasSharedStrings && this.cellHasSharedStringType)) {
      this.failed = true
      return false
    }
    if (!selfClosing && this.cellHasInlineStringType && cellContentHasRichTextRun(this.buffer, contentStart, contentNextIndex)) {
      this.metadataKeys.add('richTextArtifacts')
    }
    if (this.cellHasMetadataReference) {
      if (this.options.allowUnsupportedFeaturesForMetrics !== true) {
        this.failed = true
        return false
      }
      this.metadataKeys.add('cellMetadataRefs')
    }
    const row = packedAddressRow(this.cellPackedAddress)
    const column = packedAddressColumn(this.cellPackedAddress)
    if (this.cellStyleIndex !== null) {
      this.styleIndexes.add(this.cellStyleIndex)
    }
    if (this.cellHasSharedStringType) {
      this.usesSharedStrings = true
    }
    this.currentRow = row
    this.nextImplicitColumn = column + 1
    const hasValue = (contentFlags & cellContentHasValue) !== 0
    const hasFormula = (contentFlags & cellContentHasFormula) !== 0
    if (hasValue || hasFormula) {
      this.cellCount += 1
      this.rowCount = Math.max(this.rowCount, row + 1)
      this.columnCount = Math.max(this.columnCount, column + 1)
      this.minRow = Math.min(this.minRow, row)
      this.minColumn = Math.min(this.minColumn, column)
      this.maxRow = Math.max(this.maxRow, row)
      this.maxColumn = Math.max(this.maxColumn, column)
      if (hasValue) {
        this.valueCellCount += 1
      }
      if (hasFormula) {
        this.formulaCellCount += 1
      }
    }
    this.index = contentNextIndex
    return true
  }

  private startDataValidations(tagEnd: number, final: boolean): boolean {
    this.metadataKeys.add('validations')
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      this.index = tagEnd + 1
      return true
    }
    this.activeDataValidations = true
    this.index = tagEnd + 1
    if (!this.processActiveDataValidations(final)) {
      if (final) {
        this.failed = true
      }
      return false
    }
    return true
  }

  private readCellContentSummary(startIndex: number): boolean {
    let flags = 0
    let index = startIndex
    while (index < this.buffer.byteLength) {
      if (this.buffer[index] !== lessThan) {
        index += 1
        continue
      }
      const next = this.buffer[index + 1]
      if (next === slash && this.buffer[index + 2] === 99 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        const tagEnd = findTagEnd(this.buffer, index + 3)
        if (tagEnd === null) {
          return false
        }
        this.cellContentFlags = flags
        this.cellContentNextIndex = tagEnd + 1
        return true
      }
      if (next === slash && this.buffer[index + 2] === 118 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        index += 3
        continue
      }
      if (next === slash && this.buffer[index + 2] === 102 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        index += 3
        continue
      }
      if (
        next === slash &&
        this.buffer[index + 2] === 105 &&
        this.buffer[index + 3] === 115 &&
        !isXmlNameByte(this.buffer[index + 4] ?? 0)
      ) {
        index += 4
        continue
      }
      if (next === 118 && !isXmlNameByte(this.buffer[index + 2] ?? 0)) {
        flags |= cellContentHasValue
        index += 2
        continue
      }
      if (next === 102 && !isXmlNameByte(this.buffer[index + 2] ?? 0)) {
        flags |= cellContentHasFormula
        index += 2
        continue
      }
      if (next === 105 && this.buffer[index + 2] === 115 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        flags |= cellContentHasValue
        index += 3
        continue
      }
      const closing = this.buffer[index + 1] === slash
      const tag = readXmlTagName(this.buffer, index + (closing ? 2 : 1))
      if (closing && tag?.localName === 'c') {
        const tagEnd = findTagEnd(this.buffer, tag.endIndex)
        if (tagEnd === null) {
          return false
        }
        this.cellContentFlags = flags
        this.cellContentNextIndex = tagEnd + 1
        return true
      }
      if (tag?.localName === 'v' || tag?.localName === 'is') {
        flags |= cellContentHasValue
      } else if (tag?.localName === 'f') {
        flags |= cellContentHasFormula
      }
      index = tag?.endIndex ?? index + 1
    }
    return false
  }

  private readCellTagAttributes(startIndex: number, tagEnd: number): boolean {
    let index = startIndex
    let packedAddress: number | null = null
    let styleIndex: number | null = null
    let hasSharedStringType = false
    let hasInlineStringType = false
    let hasMetadataReference = false
    while (index < tagEnd) {
      while (index < tagEnd && isAsciiWhitespace(this.buffer[index] ?? 0)) {
        index += 1
      }
      const nameStart = index
      while (index < tagEnd && isXmlNameByte(this.buffer[index] ?? 0)) {
        index += 1
      }
      const nameEnd = index
      index = skipAsciiWhitespace(this.buffer, index, tagEnd)
      if (this.buffer[index] !== 61) {
        index += 1
        continue
      }
      index = skipAsciiWhitespace(this.buffer, index + 1, tagEnd)
      const quote = this.buffer[index]
      if (quote !== doubleQuote && quote !== singleQuote) {
        index += 1
        continue
      }
      const valueStart = index + 1
      index = valueStart
      while (index < tagEnd && this.buffer[index] !== quote) {
        index += 1
      }
      const valueEnd = index
      if (attributeNameMatches(this.buffer, nameStart, nameEnd, 'r')) {
        packedAddress = decodePackedCellAddressBytes(this.buffer, valueStart, valueEnd)
        if (packedAddress === null) {
          return false
        }
      } else if (attributeNameMatches(this.buffer, nameStart, nameEnd, 's')) {
        styleIndex = readNonNegativeIntegerFromRange(this.buffer, valueStart, valueEnd)
        if (styleIndex === null) {
          return false
        }
      } else if (attributeNameMatches(this.buffer, nameStart, nameEnd, 't')) {
        hasSharedStringType = valueEnd - valueStart === 1 && this.buffer[valueStart] === 115
        hasInlineStringType = attributeValueMatches(this.buffer, valueStart, valueEnd, 'inlineStr')
      } else if (
        nameEnd - nameStart === 2 &&
        ((this.buffer[nameStart] === 99 && this.buffer[nameStart + 1] === 109) ||
          (this.buffer[nameStart] === 118 && this.buffer[nameStart + 1] === 109))
      ) {
        hasMetadataReference = true
      }
      index += 1
    }
    this.cellPackedAddress = packedAddress ?? this.readImplicitCellPackedAddress()
    this.cellStyleIndex = styleIndex
    this.cellHasSharedStringType = hasSharedStringType
    this.cellHasInlineStringType = hasInlineStringType
    this.cellHasMetadataReference = hasMetadataReference
    return this.cellPackedAddress !== -1
  }

  private readKnownOpeningTagCode(tagStart: number): number {
    const nameStart = tagStart + 1
    switch (this.buffer[nameStart] ?? 0) {
      case 99:
        if (!isXmlNameByte(this.buffer[nameStart + 1] ?? 0)) {
          this.knownTagNameEnd = nameStart + 1
          return headlessTagCell
        }
        break
      case 100:
        if (matchesDimensionTagName(this.buffer, nameStart)) {
          this.knownTagNameEnd = nameStart + 9
          return headlessTagDimension
        }
        break
      case 114:
        if (this.buffer[nameStart + 1] === 111 && this.buffer[nameStart + 2] === 119 && !isXmlNameByte(this.buffer[nameStart + 3] ?? 0)) {
          this.knownTagNameEnd = nameStart + 3
          return headlessTagRow
        }
        break
      default:
        break
    }
    this.knownTagNameEnd = 0
    return headlessTagUnknown
  }

  private readImplicitCellPackedAddress(): number {
    return this.currentRow < 0 || this.nextImplicitColumn >= packedAddressColumnFactor
      ? -1
      : packCellAddress(this.currentRow, this.nextImplicitColumn)
  }

  private countMetadataElement(localName: string, nameEnd: number, tagEnd: number, final: boolean): boolean {
    this.recordMetadataKeys(localName)
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      this.finalizeMetadataCount({
        localName,
        childName: metadataChildName(localName),
        multiplier: metadataCountMultiplier(localName, this.buffer, nameEnd, tagEnd),
        childCount: 0,
      })
      this.index = tagEnd + 1
      return true
    }
    this.activeMetadata = {
      localName,
      childName: metadataChildName(localName),
      multiplier: metadataCountMultiplier(localName, this.buffer, nameEnd, tagEnd),
      childCount: 0,
    }
    this.index = tagEnd + 1
    if (!this.processActiveMetadata(final)) {
      if (final) {
        this.failed = true
      }
      return false
    }
    return true
  }

  private processActiveMetadata(final: boolean): boolean {
    const active = this.activeMetadata
    if (active === null) {
      return true
    }
    while (this.index < this.buffer.byteLength) {
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const closing = this.buffer[this.index + 1] === slash
      const tagNameStart = this.index + (closing ? 2 : 1)
      const tag = readXmlTagName(this.buffer, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= this.buffer.byteLength) {
          return false
        }
        this.index += 1
        continue
      }
      const tagEnd = findTagEnd(this.buffer, tag.endIndex)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return false
      }
      if (closing && tag.localName === active.localName) {
        this.finalizeMetadataCount(active)
        this.activeMetadata = null
        this.index = tagEnd + 1
        return true
      }
      if (!closing && active.childName !== null && tag.localName === active.childName) {
        active.childCount += 1
      }
      this.index = tagEnd + 1
    }
    if (final) {
      this.failed = true
    } else {
      this.index = this.buffer.byteLength
    }
    return false
  }

  private processActiveDataValidations(final: boolean): boolean {
    while (!this.failed && this.index < this.buffer.byteLength) {
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const closing = this.buffer[this.index + 1] === slash
      const tagNameStart = this.index + (closing ? 2 : 1)
      const tag = readXmlTagName(this.buffer, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= this.buffer.byteLength) {
          return false
        }
        this.index += 1
        continue
      }
      const tagEnd = findTagEnd(this.buffer, tag.endIndex)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return false
      }
      if (closing && tag.localName === 'dataValidations') {
        this.activeDataValidations = false
        this.index = tagEnd + 1
        return true
      }
      if (!closing && tag.localName === 'dataValidation') {
        if (!this.processActiveDataValidationElement(tagEnd, final)) {
          return false
        }
        continue
      }
      this.index = tagEnd + 1
    }
    if (final) {
      this.failed = true
    } else {
      this.index = this.buffer.byteLength
    }
    return false
  }

  private processActiveDataValidationElement(tagEnd: number, final: boolean): boolean {
    const startIndex = this.index
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(this.buffer, contentStart, 'dataValidation')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      this.index = startIndex
      return false
    }
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    const count = countLargeSimpleDataValidationsFromBytes('Sheet1', this.buffer, startIndex, endIndex)
    if (count === null) {
      if (this.options.allowUnsupportedFeaturesForMetrics === true) {
        this.index = endIndex
        return true
      }
      this.failed = true
      return true
    }
    this.dataValidationCount += count
    this.index = endIndex
    return true
  }

  private finalizeMetadataCount(active: ActiveMetadataCount): void {
    if (active.localName === 'conditionalFormatting') {
      this.conditionalFormatCount += active.multiplier * Math.max(1, active.childCount)
    } else if (active.localName === 'mergeCells') {
      this.mergeCount += active.childCount
    } else if (active.localName === 'tableParts') {
      this.tableCount += active.childCount
    }
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.byteLength)
  }

  private recordMetadataKeys(localName: string): void {
    for (const key of sheetMetadataKeysForHeadlessElement(localName)) {
      this.metadataKeys.add(key)
    }
  }
}
