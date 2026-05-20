import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
} from '@bilig/protocol'
import { readLargeSimpleAutoFiltersFromBytes } from './xlsx-large-simple-autofilter-byte-scan.js'
import { readLargeSimpleConditionalFormattingFromBytes } from './xlsx-large-simple-conditional-format-byte-scan.js'
import {
  countLargeSimpleDataValidationsFromBytes,
  readLargeSimpleDataValidationsFromBytes,
} from './xlsx-large-simple-data-validation-byte-scan.js'
import { readLargeSimpleSheetHyperlinkRefsFromBytes } from './xlsx-large-simple-hyperlinks.js'
import {
  appendLargeSimpleColumnMetadataFromBytes,
  appendLargeSimpleRowMetadataTagFromBytes,
  readLargeSimpleDrawingRelationshipIdTagFromBytes,
  readLargeSimpleMergeRefsFromBytes,
  readLargeSimpleSheetFormatPrTagFromBytes,
  readLargeSimpleTableRelationshipIdsFromBytes,
} from './xlsx-large-simple-metadata-byte-scan.js'
import { appendLargeSimplePrintPageSetupElement, isLargeSimplePrintPageSetupElementName } from './xlsx-large-simple-printer-settings.js'
import { rowTagHasMetadataAttribute } from './xlsx-large-simple-row-metadata-scan.js'
import type {
  LargeSimpleWorksheetCellMetadataRef,
  LargeSimpleWorksheetMergeRef,
  LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import { decodeBytes } from './xlsx-large-simple-xml-byte-utils.js'
import {
  countOpeningTags,
  findClosingTag,
  findTagEnd,
  isSelfClosingTag,
  readXmlAttributeFromTag,
  readXmlAttributeRangeFromTag,
  readXmlTagName,
} from './xlsx-large-simple-worksheet-stream-xml.js'

const lessThan = 60
const slash = 47
const conditionalFormattingCloseBytes = new Uint8Array([
  60, 47, 99, 111, 110, 100, 105, 116, 105, 111, 110, 97, 108, 70, 111, 114, 109, 97, 116, 116, 105, 110, 103, 62,
])
const extensionElementPattern = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?ext>)/gu
const slicerListElementPattern = /<(?:[A-Za-z_][\w.-]*:)?slicerList\b/u

type StreamedMetadataElement = 'mergeCells' | 'tableParts'

interface ActiveConditionalFormatting {
  readonly rootTag: Uint8Array
  ruleSeen: boolean
}

export interface LargeSimpleWorksheetMetadataProcessResult {
  readonly index: number
  readonly complete: boolean
  readonly failed: boolean
}

export class LargeSimpleWorksheetStreamMetadataScanner {
  private worksheetRootOpenTag: string | undefined
  private columnEntries: WorkbookAxisEntrySnapshot[] | undefined
  private columnMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private conditionalFormats: WorkbookConditionalFormatSnapshot[] | undefined
  private conditionalFormatIdCounter = 0
  private conditionalFormattingXml: string[] | undefined
  private dataValidations: WorkbookDataValidationSnapshot[] | undefined
  private controlArtifactsXml: string[] | undefined
  private legacyDrawingRelationshipId: string | undefined
  private cellMetadataRefs: LargeSimpleWorksheetCellMetadataRef[] | undefined
  private drawingRelationshipId: string | undefined
  private filters: LargeSimpleWorksheetScannedMetadata['filters']
  private hyperlinks: LargeSimpleWorksheetScannedMetadata['hyperlinks']
  private rowEntries: WorkbookAxisEntrySnapshot[] | undefined
  private rowMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private mergeRefs: LargeSimpleWorksheetMergeRef[] | undefined
  private printPageSetup: LargeSimpleWorksheetScannedMetadata['printPageSetup']
  private sheetFormatPr: LargeSimpleWorksheetScannedMetadata['sheetFormatPr']
  private sheetSlicerListExtXml: string | undefined
  private tableRelationshipIds: string[] | undefined
  private activeMetadataElement: StreamedMetadataElement | null = null
  private activeConditionalFormatting: ActiveConditionalFormatting | null = null
  private activeDataValidations = false
  private mergeCountValue = 0
  private conditionalFormatCountValue = 0
  private dataValidationCountValue = 0
  private tableCountValue = 0

  constructor(
    private readonly options: {
      readonly retainMetadataXml: boolean
      readonly sheetName: string | undefined
    },
  ) {}

  get mergeCount(): number {
    return this.mergeCountValue
  }

  get conditionalFormatCount(): number {
    return this.conditionalFormatCountValue
  }

  get dataValidationCount(): number {
    return this.dataValidationCountValue
  }

  get tableCount(): number {
    return this.tableCountValue
  }

  hasActiveStream(): boolean {
    return this.activeMetadataElement !== null || this.activeConditionalFormatting !== null || this.activeDataValidations
  }

  collectWorksheetRootOpenTag(bytes: Uint8Array, startIndex: number, endIndex: number): void {
    if (this.options.retainMetadataXml) {
      this.worksheetRootOpenTag = decodeBytes(bytes, startIndex, endIndex)
    }
  }

  collectRowMetadata(bytes: Uint8Array, nameEnd: number, tagEnd: number, currentRow: number): void {
    if (!this.options.retainMetadataXml || !rowTagHasMetadataAttribute(bytes, nameEnd, tagEnd)) {
      return
    }
    this.rowEntries ??= []
    this.rowMetadata ??= []
    appendLargeSimpleRowMetadataTagFromBytes(this.rowEntries, this.rowMetadata, bytes, nameEnd, tagEnd, currentRow)
  }

  collectCellMetadataRef(bytes: Uint8Array, nameEnd: number, tagEnd: number, address: string): void {
    if (!this.options.retainMetadataXml) {
      return
    }
    const cm = readXmlAttributeFromTag(bytes, nameEnd, tagEnd, 'cm')
    const vm = readXmlAttributeFromTag(bytes, nameEnd, tagEnd, 'vm')
    if (!cm && !vm) {
      return
    }
    this.cellMetadataRefs ??= []
    this.cellMetadataRefs.push({
      address,
      ...(cm ? { cm } : {}),
      ...(vm ? { vm } : {}),
    })
  }

  buildMetadataScan(): LargeSimpleWorksheetScannedMetadata | undefined {
    const columns =
      (this.columnEntries?.length ?? 0) > 0 || (this.columnMetadata?.length ?? 0) > 0
        ? { entries: this.columnEntries ?? [], metadata: this.columnMetadata ?? [] }
        : undefined
    const rows =
      (this.rowEntries?.length ?? 0) > 0 || (this.rowMetadata?.length ?? 0) > 0
        ? { entries: this.rowEntries ?? [], metadata: this.rowMetadata ?? [] }
        : undefined
    const metadata: LargeSimpleWorksheetScannedMetadata = {
      ...(this.cellMetadataRefs && this.cellMetadataRefs.length > 0 ? { cellMetadataRefs: this.cellMetadataRefs } : {}),
      ...(columns ? { columns } : {}),
      ...(this.conditionalFormats && this.conditionalFormats.length > 0 ? { conditionalFormats: this.conditionalFormats } : {}),
      ...(this.conditionalFormattingXml && this.conditionalFormattingXml.length > 0
        ? { conditionalFormattingXml: this.conditionalFormattingXml }
        : {}),
      ...(this.controlArtifactsXml && this.controlArtifactsXml.length > 0 && this.worksheetRootOpenTag
        ? {
            controlArtifacts: {
              controlsXml: this.controlArtifactsXml.join(''),
              worksheetRootOpenTag: this.worksheetRootOpenTag,
              ...(this.legacyDrawingRelationshipId ? { legacyDrawingRelationshipId: this.legacyDrawingRelationshipId } : {}),
            },
          }
        : {}),
      ...(this.dataValidations && this.dataValidations.length > 0 ? { dataValidations: this.dataValidations } : {}),
      ...(this.drawingRelationshipId ? { drawingRelationshipId: this.drawingRelationshipId } : {}),
      ...(this.legacyDrawingRelationshipId ? { legacyDrawingRelationshipId: this.legacyDrawingRelationshipId } : {}),
      ...(this.filters && this.filters.length > 0 ? { filters: this.filters } : {}),
      ...(this.hyperlinks && this.hyperlinks.length > 0 ? { hyperlinks: this.hyperlinks } : {}),
      ...(rows ? { rows } : {}),
      ...(this.mergeRefs && this.mergeRefs.length > 0 ? { merges: this.mergeRefs } : {}),
      ...(this.printPageSetup ? { printPageSetup: this.printPageSetup } : {}),
      ...(this.sheetFormatPr ? { sheetFormatPr: this.sheetFormatPr } : {}),
      ...(this.sheetSlicerListExtXml ? { sheetSlicerListExtXml: this.sheetSlicerListExtXml } : {}),
      ...(this.tableRelationshipIds && this.tableRelationshipIds.length > 0 ? { tableRelationshipIds: this.tableRelationshipIds } : {}),
    }
    return Object.keys(metadata).length > 0 ? metadata : undefined
  }

  collectMetadataElement(
    localName: string,
    bytes: Uint8Array,
    startIndex: number,
    tagEnd: number,
    final: boolean,
  ): LargeSimpleWorksheetMetadataProcessResult {
    if (localName === 'dataValidations' && isSelfClosingTag(bytes, tagEnd)) {
      return completed(tagEnd + 1)
    }
    if (isSelfClosingTag(bytes, tagEnd)) {
      const handled = this.options.retainMetadataXml && this.collectTypedMetadataElement(localName, bytes, startIndex, tagEnd + 1)
      if (!handled) {
        this.countMetadataElement(localName, bytes, tagEnd + 1, tagEnd + 1)
      }
      return this.options.retainMetadataXml && !handled ? failed(tagEnd + 1) : completed(tagEnd + 1)
    }
    if (localName === 'conditionalFormatting') {
      this.activeConditionalFormatting = {
        rootTag: bytes.slice(startIndex, tagEnd + 1),
        ruleSeen: false,
      }
      return this.processActiveConditionalFormatting(bytes, tagEnd + 1, final)
    }
    if (localName === 'mergeCells' || localName === 'tableParts') {
      this.activeMetadataElement = localName
      return this.processActiveMetadataElement(bytes, tagEnd + 1, final)
    }
    if (localName === 'dataValidations') {
      this.activeDataValidations = true
      return this.processActiveDataValidations(bytes, tagEnd + 1, final)
    }
    const closing = findClosingTag(bytes, tagEnd + 1, localName)
    if (!closing) {
      return final ? failed(startIndex, false) : incomplete(startIndex)
    }
    const handled = this.options.retainMetadataXml && this.collectTypedMetadataElement(localName, bytes, startIndex, closing.end)
    if (!handled) {
      this.countMetadataElement(localName, bytes, tagEnd + 1, closing.start)
    }
    return this.options.retainMetadataXml && !handled ? failed(closing.end) : completed(closing.end)
  }

  processActive(bytes: Uint8Array, index: number, final: boolean): LargeSimpleWorksheetMetadataProcessResult {
    if (this.activeMetadataElement !== null) {
      return this.processActiveMetadataElement(bytes, index, final)
    }
    if (this.activeConditionalFormatting !== null) {
      return this.processActiveConditionalFormatting(bytes, index, final)
    }
    if (this.activeDataValidations) {
      return this.processActiveDataValidations(bytes, index, final)
    }
    return completed(index)
  }

  private collectTypedMetadataElement(localName: string, bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
    if (localName === 'mergeCells') {
      const refs = readLargeSimpleMergeRefsFromBytes(bytes, startIndex, endIndex)
      this.mergeCountValue += refs.length
      if (refs.length > 0) {
        this.mergeRefs ??= []
        this.mergeRefs.push(...refs)
      }
      return true
    }
    if (localName === 'cols') {
      this.columnEntries ??= []
      this.columnMetadata ??= []
      appendLargeSimpleColumnMetadataFromBytes(this.columnEntries, this.columnMetadata, bytes, startIndex, endIndex)
      return true
    }
    if (localName === 'sheetFormatPr') {
      this.sheetFormatPr = readLargeSimpleSheetFormatPrTagFromBytes(bytes, startIndex, endIndex) ?? this.sheetFormatPr
      return true
    }
    if (localName === 'extLst') {
      const sheetSlicerListExtXml = readSlicerListExtensionXml(decodeBytes(bytes, startIndex, endIndex))
      if (!sheetSlicerListExtXml) {
        return false
      }
      this.sheetSlicerListExtXml = sheetSlicerListExtXml
      return true
    }
    if (localName === 'drawing') {
      this.drawingRelationshipId = readLargeSimpleDrawingRelationshipIdTagFromBytes(bytes, startIndex, endIndex)
      return true
    }
    if (localName === 'legacyDrawing') {
      const tagXml = decodeBytes(bytes, startIndex, endIndex)
      this.legacyDrawingRelationshipId =
        readElementAttribute(tagXml, 'r:id') ?? readElementAttribute(tagXml, 'id') ?? this.legacyDrawingRelationshipId
      return true
    }
    if (localName === 'oleObjects') {
      this.controlArtifactsXml ??= []
      this.controlArtifactsXml.push(decodeBytes(bytes, startIndex, endIndex))
      return true
    }
    if (localName === 'autoFilter') {
      if (!this.options.sheetName) {
        return false
      }
      const filters = readLargeSimpleAutoFiltersFromBytes(this.options.sheetName, bytes, startIndex, endIndex)
      this.filters = [...(this.filters ?? []), ...filters]
      return true
    }
    if (localName === 'hyperlinks') {
      const refs = readLargeSimpleSheetHyperlinkRefsFromBytes(bytes, startIndex, endIndex)
      if (refs === null) {
        return false
      }
      this.hyperlinks = [...(this.hyperlinks ?? []), ...refs]
      return true
    }
    if (isLargeSimplePrintPageSetupElementName(localName)) {
      this.printPageSetup ??= {}
      appendLargeSimplePrintPageSetupElement(this.printPageSetup, localName, decodeBytes(bytes, startIndex, endIndex))
      return true
    }
    if (localName === 'tableParts') {
      const relationshipIds = readLargeSimpleTableRelationshipIdsFromBytes(bytes, startIndex, endIndex)
      this.tableCountValue += relationshipIds.length
      if (relationshipIds.length > 0) {
        this.tableRelationshipIds ??= []
        this.tableRelationshipIds.push(...relationshipIds)
      }
      return true
    }
    if (localName === 'conditionalFormatting') {
      if (!this.options.sheetName) {
        return false
      }
      const scan = readLargeSimpleConditionalFormattingFromBytes(
        this.options.sheetName,
        bytes,
        startIndex,
        endIndex,
        this.conditionalFormatIdCounter + 1,
      )
      this.conditionalFormatCountValue += scan.ruleCount
      if (scan.conditionalFormats && scan.conditionalFormats.length > 0) {
        this.conditionalFormats ??= []
        this.conditionalFormats.push(...scan.conditionalFormats)
        this.conditionalFormatIdCounter += scan.conditionalFormats.length
      }
      if (scan.artifactXml) {
        this.conditionalFormattingXml ??= []
        this.conditionalFormattingXml.push(scan.artifactXml)
      }
      return true
    }
    return false
  }

  private countMetadataElement(localName: string, bytes: Uint8Array, contentStart: number, contentEnd: number): void {
    if (localName === 'conditionalFormatting') {
      this.conditionalFormatCountValue += Math.max(1, countOpeningTags(bytes, contentStart, contentEnd, 'cfRule'))
      return
    }
    if (localName === 'mergeCells') {
      this.mergeCountValue += countOpeningTags(bytes, contentStart, contentEnd, 'mergeCell')
    } else if (localName === 'tableParts') {
      this.tableCountValue += countOpeningTags(bytes, contentStart, contentEnd, 'tablePart')
    } else if (localName === 'dataValidations') {
      this.dataValidationCountValue += countOpeningTags(bytes, contentStart, contentEnd, 'dataValidation')
    }
  }

  private processActiveMetadataElement(bytes: Uint8Array, index: number, final: boolean): LargeSimpleWorksheetMetadataProcessResult {
    const activeElement = this.activeMetadataElement
    if (activeElement === null) {
      return completed(index)
    }
    while (index < bytes.byteLength) {
      if (bytes[index] !== lessThan) {
        index += 1
        continue
      }
      const closing = bytes[index + 1] === slash
      const tagNameStart = index + (closing ? 2 : 1)
      const tag = readXmlTagName(bytes, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= bytes.byteLength) {
          return incomplete(index)
        }
        index += 1
        continue
      }
      const tagEnd = findTagEnd(bytes, tag.endIndex)
      if (tagEnd === null) {
        return final ? failed(index, false) : incomplete(index)
      }
      if (closing && tag.localName === activeElement) {
        this.activeMetadataElement = null
        return completed(tagEnd + 1)
      }
      if (!closing) {
        this.collectActiveMetadataTag(activeElement, tag.localName, bytes, tag.endIndex, tagEnd)
      }
      index = tagEnd + 1
    }
    return final ? failed(index, false) : incomplete(bytes.byteLength)
  }

  private processActiveConditionalFormatting(bytes: Uint8Array, index: number, final: boolean): LargeSimpleWorksheetMetadataProcessResult {
    const active = this.activeConditionalFormatting
    if (active === null) {
      return completed(index)
    }
    while (index < bytes.byteLength) {
      if (bytes[index] !== lessThan) {
        index += 1
        continue
      }
      const closing = bytes[index + 1] === slash
      const tagNameStart = index + (closing ? 2 : 1)
      const tag = readXmlTagName(bytes, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= bytes.byteLength) {
          return incomplete(index)
        }
        index += 1
        continue
      }
      const tagEnd = findTagEnd(bytes, tag.endIndex)
      if (tagEnd === null) {
        return final ? failed(index, false) : incomplete(index)
      }
      if (closing && tag.localName === 'conditionalFormatting') {
        if (!active.ruleSeen) {
          this.conditionalFormatCountValue += 1
        }
        this.activeConditionalFormatting = null
        return completed(tagEnd + 1)
      }
      if (!closing && tag.localName === 'cfRule') {
        const result = this.processActiveConditionalFormatRule(active, bytes, index, tagEnd, final)
        if (!result.complete || result.failed) {
          return result
        }
        index = result.index
        continue
      }
      if (!closing && this.options.retainMetadataXml) {
        return failed(index)
      }
      index = tagEnd + 1
    }
    return final ? failed(index, false) : incomplete(bytes.byteLength)
  }

  private processActiveConditionalFormatRule(
    active: ActiveConditionalFormatting,
    bytes: Uint8Array,
    startIndex: number,
    tagEnd: number,
    final: boolean,
  ): LargeSimpleWorksheetMetadataProcessResult {
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(bytes, contentStart, 'cfRule')
    if (!closing) {
      return final ? failed(startIndex, false) : incomplete(startIndex)
    }
    active.ruleSeen = true
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    if (!this.options.retainMetadataXml) {
      this.conditionalFormatCountValue += activeConditionalFormattingRangeCount(active)
      return completed(endIndex)
    }
    if (!this.options.sheetName) {
      return failed(startIndex)
    }
    const scanBytes = wrapConditionalFormatRule(active.rootTag, bytes, startIndex, endIndex)
    const scan = readLargeSimpleConditionalFormattingFromBytes(
      this.options.sheetName,
      scanBytes,
      0,
      scanBytes.byteLength,
      this.conditionalFormatIdCounter + 1,
    )
    this.conditionalFormatCountValue += scan.ruleCount
    if (scan.conditionalFormats && scan.conditionalFormats.length > 0) {
      this.conditionalFormats ??= []
      this.conditionalFormats.push(...scan.conditionalFormats)
      this.conditionalFormatIdCounter += scan.conditionalFormats.length
    }
    if (scan.artifactXml) {
      this.conditionalFormattingXml ??= []
      this.conditionalFormattingXml.push(scan.artifactXml)
    }
    return completed(endIndex)
  }

  private processActiveDataValidations(bytes: Uint8Array, index: number, final: boolean): LargeSimpleWorksheetMetadataProcessResult {
    while (index < bytes.byteLength) {
      if (bytes[index] !== lessThan) {
        index += 1
        continue
      }
      const closing = bytes[index + 1] === slash
      const tagNameStart = index + (closing ? 2 : 1)
      const tag = readXmlTagName(bytes, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= bytes.byteLength) {
          return incomplete(index)
        }
        index += 1
        continue
      }
      const tagEnd = findTagEnd(bytes, tag.endIndex)
      if (tagEnd === null) {
        return final ? failed(index, false) : incomplete(index)
      }
      if (closing && tag.localName === 'dataValidations') {
        this.activeDataValidations = false
        return completed(tagEnd + 1)
      }
      if (!closing && tag.localName === 'dataValidation') {
        const result = this.processActiveDataValidationElement(bytes, index, tagEnd, final)
        if (!result.complete || result.failed) {
          return result
        }
        index = result.index
        continue
      }
      index = tagEnd + 1
    }
    return final ? failed(index, false) : incomplete(bytes.byteLength)
  }

  private processActiveDataValidationElement(
    bytes: Uint8Array,
    startIndex: number,
    tagEnd: number,
    final: boolean,
  ): LargeSimpleWorksheetMetadataProcessResult {
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(bytes, contentStart, 'dataValidation')
    if (!closing) {
      return final ? failed(startIndex, false) : incomplete(startIndex)
    }
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    if (this.options.retainMetadataXml) {
      if (!this.options.sheetName) {
        return failed(startIndex)
      }
      const validations = readLargeSimpleDataValidationsFromBytes(this.options.sheetName, bytes, startIndex, endIndex)
      if (validations === null) {
        return failed(startIndex)
      }
      this.dataValidationCountValue += validations.length
      if (validations.length > 0) {
        this.dataValidations ??= []
        this.dataValidations.push(...validations)
      }
    } else {
      const count = countLargeSimpleDataValidationsFromBytes(this.options.sheetName ?? 'Sheet1', bytes, startIndex, endIndex)
      if (count === null) {
        return failed(startIndex)
      }
      this.dataValidationCountValue += count
    }
    return completed(endIndex)
  }

  private collectActiveMetadataTag(
    activeElement: StreamedMetadataElement,
    localName: string,
    bytes: Uint8Array,
    nameEnd: number,
    tagEnd: number,
  ): void {
    if (activeElement === 'mergeCells' && localName === 'mergeCell') {
      this.collectMergeCellTag(bytes, nameEnd, tagEnd)
      return
    }
    if (activeElement === 'tableParts' && localName === 'tablePart') {
      const relationshipId =
        readXmlAttributeFromTag(bytes, nameEnd, tagEnd, 'r:id') ?? readXmlAttributeFromTag(bytes, nameEnd, tagEnd, 'id')
      if (relationshipId) {
        this.tableCountValue += 1
        if (this.options.retainMetadataXml) {
          this.tableRelationshipIds ??= []
          this.tableRelationshipIds.push(relationshipId)
        }
      }
    }
  }

  private collectMergeCellTag(bytes: Uint8Array, nameEnd: number, tagEnd: number): void {
    const ref = readXmlAttributeFromTag(bytes, nameEnd, tagEnd, 'ref')
    const [startAddress, endAddress] = ref?.split(':') ?? []
    if (!startAddress || !endAddress || startAddress === endAddress) {
      return
    }
    this.mergeCountValue += 1
    if (this.options.retainMetadataXml) {
      this.mergeRefs ??= []
      this.mergeRefs.push({ startAddress, endAddress })
    }
  }
}

function completed(index: number): LargeSimpleWorksheetMetadataProcessResult {
  return { index, complete: true, failed: false }
}

function incomplete(index: number): LargeSimpleWorksheetMetadataProcessResult {
  return { index, complete: false, failed: false }
}

function failed(index: number, complete = true): LargeSimpleWorksheetMetadataProcessResult {
  return { index, complete, failed: true }
}

function wrapConditionalFormatRule(rootTag: Uint8Array, bytes: Uint8Array, startIndex: number, endIndex: number): Uint8Array {
  const ruleBytes = bytes.subarray(startIndex, endIndex)
  const output = new Uint8Array(rootTag.byteLength + ruleBytes.byteLength + conditionalFormattingCloseBytes.byteLength)
  output.set(rootTag)
  output.set(ruleBytes, rootTag.byteLength)
  output.set(conditionalFormattingCloseBytes, rootTag.byteLength + ruleBytes.byteLength)
  return output
}

function activeConditionalFormattingRangeCount(active: ActiveConditionalFormatting): number {
  const tag = readXmlTagName(active.rootTag, 1)
  return tag ? countSqrefRangesFromTag(active.rootTag, tag.endIndex, active.rootTag.byteLength - 1) : 1
}

function readSlicerListExtensionXml(xml: string): string | undefined {
  extensionElementPattern.lastIndex = 0
  return [...xml.matchAll(extensionElementPattern)].find((match) => slicerListElementPattern.test(match[0]))?.[0]
}

function readElementAttribute(xml: string, name: string): string | null {
  return new RegExp(`\\s${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function countSqrefRangesFromTag(bytes: Uint8Array, nameEnd: number, tagEnd: number): number {
  const sqref = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, 'sqref')
  if (!sqref) {
    return 1
  }
  let count = 0
  let inToken = false
  for (let index = sqref.start; index < sqref.end; index += 1) {
    if (isWhitespaceByte(bytes[index] ?? 0)) {
      inToken = false
      continue
    }
    if (!inToken) {
      count += 1
      inToken = true
    }
  }
  return Math.max(1, count)
}

function isWhitespaceByte(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}
