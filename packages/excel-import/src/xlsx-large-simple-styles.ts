import type {
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleFontSnapshot,
  CellStyleProtectionSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
} from '@bilig/protocol'
import { builtinNumberFormatCode, LargeSimpleNumberFormatCollector } from './xlsx-large-simple-number-formats.js'
import {
  asciiByteRangeEquals,
  closingStringTagRetainLength,
  findByteInRange,
  findByteTagEnd,
  findClosingStringElementEnd,
  findNextOpeningTag,
  findNextParentBoundaryOrChild,
  findStringTagEnd,
  inspectRequiredXfOpeningTag,
  isApplied,
  isSelfClosingByteTag,
  isSelfClosingStringTag,
  isXmlNameByte,
  readAttribute,
  readBooleanAttribute,
  readNonNegativeIntegerAttribute,
  readNumberAttribute,
} from './xlsx-large-simple-style-xml-utils.js'

type ImportedCellStyle = Omit<CellStyleRecord, 'id'>

const elementTextCache = new Map<string, RegExp>()
const lessThan = 60
const slash = 47
const styleSupportTagUnknown = 0
const styleSupportTagCellXfs = 1
const styleSupportTagXf = 2
const styleSupportTagOther = 3
const xfSupportUnsupported = 1 << 0
const xfSupportPotentialVisual = 1 << 1
const emptyBytes = new Uint8Array(0)

export interface LargeSimpleWorkbookStylesScanOptions {
  readonly onRetainedBufferLength?: (length: number) => void
}

type LargeSimpleChunkConsumer = (chunk: Uint8Array) => boolean | void

export interface LargeSimpleWorkbookStyleArtifacts {
  readonly stylesByIndex: Map<number, ImportedCellStyle> | null
  readonly numberFormatsByStyleIndex: Map<number, string> | null
}

export interface LargeSimpleWorkbookStyleSupportScan {
  readonly hasUnsupportedStyles: boolean
  readonly hasPotentialVisualStyles: boolean
}

export function readLargeSimpleWorkbookStyles(
  stylesXml: string | null,
  requiredStyleIndexes: ReadonlySet<number>,
): Map<number, ImportedCellStyle> | null {
  if (!stylesXml || requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const fills = readFillStyles(stylesXml)
  const fonts = readFontStyles(stylesXml)
  const cellXfsXml = extractElementXml(stylesXml, 'cellXfs')
  if (!fills || !fonts || !cellXfsXml) {
    return null
  }
  const cellXfs = [
    ...cellXfsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?xf>)/gu),
  ]
  const styles = new Map<number, ImportedCellStyle>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs[styleIndex]?.[0]
    if (!xfXml) {
      return null
    }
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0]
    if (!openingTag) {
      return null
    }
    const fillId = readNonNegativeIntegerAttribute(openingTag, 'fillId')
    const fontId = readNonNegativeIntegerAttribute(openingTag, 'fontId')
    const borderId = readNonNegativeIntegerAttribute(openingTag, 'borderId')
    if (isApplied(openingTag, 'applyBorder', borderId)) {
      return null
    }
    const fill = isApplied(openingTag, 'applyFill', fillId) ? fills[fillId ?? -1] : undefined
    const font = isApplied(openingTag, 'applyFont', fontId) ? fonts[fontId ?? -1] : undefined
    const alignment = readBooleanAttribute(readAttribute(openingTag, 'applyAlignment')) === true ? readAlignmentStyle(xfXml) : undefined
    const protection =
      readBooleanAttribute(readAttribute(openingTag, 'applyProtection')) === true
        ? (readProtectionStyle(xfXml) ?? {})
        : readProtectionStyle(xfXml)
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
      ...(alignment ? { alignment } : {}),
      ...(protection !== undefined ? { protection } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

export function readLargeSimpleWorkbookStyleArtifactsFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): LargeSimpleWorkbookStyleArtifacts {
  if (requiredStyleIndexes.size === 0) {
    return { stylesByIndex: new Map(), numberFormatsByStyleIndex: new Map() }
  }
  const cellXfs = collectIndexedXmlElementsFromChunks(readChunks, 'cellXfs', 'xf', requiredStyleIndexes, options)
  if (!cellXfs) {
    return { stylesByIndex: null, numberFormatsByStyleIndex: null }
  }
  const requiredFillIndexes = new Set<number>()
  const requiredFontIndexes = new Set<number>()
  const requiredCustomFormatIds = new Set<number>()
  const styleRefs = new Map<number, StyleComponentRefs>()
  const styleFormatIds = new Map<number, number>()
  let stylesSupported = true
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs.get(styleIndex)
    if (!xfXml) {
      return { stylesByIndex: null, numberFormatsByStyleIndex: null }
    }
    const openingTag = readXfOpeningTag(xfXml)
    if (!openingTag) {
      return { stylesByIndex: null, numberFormatsByStyleIndex: null }
    }
    const formatId = readNonNegativeIntegerAttribute(openingTag, 'numFmtId')
    if (formatId !== null && formatId !== 0) {
      styleFormatIds.set(styleIndex, formatId)
      if (!builtinNumberFormatCode(formatId)) {
        requiredCustomFormatIds.add(formatId)
      }
    }
    if (!stylesSupported) {
      continue
    }
    const refs = readStyleComponentRefsFromOpeningTag(xfXml, openingTag)
    if (!refs) {
      stylesSupported = false
      styleRefs.clear()
      requiredFillIndexes.clear()
      requiredFontIndexes.clear()
      continue
    }
    styleRefs.set(styleIndex, refs)
    if (refs.fillApplied && refs.fillId !== null) {
      requiredFillIndexes.add(refs.fillId)
    }
    if (refs.fontApplied && refs.fontId !== null) {
      requiredFontIndexes.add(refs.fontId)
    }
  }
  if (requiredFillIndexes.size === 0 && requiredFontIndexes.size === 0 && requiredCustomFormatIds.size === 0) {
    return {
      stylesByIndex: stylesSupported ? buildStylesByIndex(styleRefs, new Map(), new Map()) : null,
      numberFormatsByStyleIndex: buildNumberFormatsByStyleIndex(styleFormatIds, new Map()),
    }
  }
  const components = collectLargeSimpleWorkbookStyleComponentsFromChunks(
    readChunks,
    stylesSupported ? requiredFillIndexes : new Set<number>(),
    stylesSupported ? requiredFontIndexes : new Set<number>(),
    requiredCustomFormatIds,
    options,
  )
  if (!components) {
    return { stylesByIndex: null, numberFormatsByStyleIndex: null }
  }
  return {
    stylesByIndex: stylesSupported ? buildStylesByIndex(styleRefs, components.fills, components.fonts) : null,
    numberFormatsByStyleIndex: buildNumberFormatsByStyleIndex(styleFormatIds, components.numberFormats),
  }
}

export function hasUnsupportedLargeSimpleWorkbookStylesFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): boolean | null {
  const scan = inspectLargeSimpleWorkbookStyleSupportFromChunks(readChunks, requiredStyleIndexes, options)
  return scan?.hasUnsupportedStyles ?? null
}

export function inspectLargeSimpleWorkbookStyleSupportFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): LargeSimpleWorkbookStyleSupportScan | null {
  if (requiredStyleIndexes.size === 0) {
    return { hasUnsupportedStyles: false, hasPotentialVisualStyles: false }
  }
  const scanner = new RequiredXfUnsupportedStyleScanner(requiredStyleIndexes, options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export async function inspectLargeSimpleWorkbookStyleSupportFromChunksAsync(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => Promise<boolean>,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): Promise<LargeSimpleWorkbookStyleSupportScan | null> {
  if (requiredStyleIndexes.size === 0) {
    return { hasUnsupportedStyles: false, hasPotentialVisualStyles: false }
  }
  const scanner = new RequiredXfUnsupportedStyleScanner(requiredStyleIndexes, options)
  if (!(await readChunks((chunk) => scanner.push(chunk)))) {
    return null
  }
  return scanner.finish()
}

export function readLargeSimpleWorkbookStylesFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  requiredStyleIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions = {},
): Map<number, ImportedCellStyle> | null {
  if (requiredStyleIndexes.size === 0) {
    return new Map()
  }
  const cellXfs = collectIndexedXmlElementsFromChunks(readChunks, 'cellXfs', 'xf', requiredStyleIndexes, options)
  if (!cellXfs) {
    return null
  }
  const requiredFillIndexes = new Set<number>()
  const requiredFontIndexes = new Set<number>()
  const styleRefs = new Map<number, StyleComponentRefs>()
  for (const styleIndex of requiredStyleIndexes) {
    const xfXml = cellXfs.get(styleIndex)
    if (!xfXml) {
      return null
    }
    const refs = readStyleComponentRefs(xfXml)
    if (!refs) {
      return null
    }
    styleRefs.set(styleIndex, refs)
    if (refs.fillApplied && refs.fillId !== null) {
      requiredFillIndexes.add(refs.fillId)
    }
    if (refs.fontApplied && refs.fontId !== null) {
      requiredFontIndexes.add(refs.fontId)
    }
  }
  const fills = collectIndexedXmlElementsFromChunks(readChunks, 'fills', 'fill', requiredFillIndexes, options)
  const fonts = collectIndexedXmlElementsFromChunks(readChunks, 'fonts', 'font', requiredFontIndexes, options)
  if (!fills || !fonts) {
    return null
  }
  const styles = new Map<number, ImportedCellStyle>()
  for (const [styleIndex, refs] of styleRefs) {
    const fill = refs.fillApplied && refs.fillId !== null ? readFillStyle(fills.get(refs.fillId) ?? '') : undefined
    const font = refs.fontApplied && refs.fontId !== null ? readFontStyle(fonts.get(refs.fontId) ?? '') : undefined
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
      ...(refs.alignment ? { alignment: refs.alignment } : {}),
      ...(refs.protection !== undefined ? { protection: refs.protection } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

interface StyleComponentRefs {
  readonly fillId: number | null
  readonly fontId: number | null
  readonly fillApplied: boolean
  readonly fontApplied: boolean
  readonly alignment?: CellStyleAlignmentSnapshot
  readonly protection?: CellStyleProtectionSnapshot
}

interface LargeSimpleWorkbookStyleComponents {
  readonly fills: Map<number, string> | null
  readonly fonts: Map<number, string> | null
  readonly numberFormats: Map<number, string> | null
}

function readStyleComponentRefs(xfXml: string): StyleComponentRefs | null {
  const openingTag = readXfOpeningTag(xfXml)
  if (!openingTag) {
    return null
  }
  return readStyleComponentRefsFromOpeningTag(xfXml, openingTag)
}

function readStyleComponentRefsFromOpeningTag(xfXml: string, openingTag: string): StyleComponentRefs | null {
  const fillId = readNonNegativeIntegerAttribute(openingTag, 'fillId')
  const fontId = readNonNegativeIntegerAttribute(openingTag, 'fontId')
  const borderId = readNonNegativeIntegerAttribute(openingTag, 'borderId')
  if (isApplied(openingTag, 'applyBorder', borderId)) {
    return null
  }
  const alignment = readBooleanAttribute(readAttribute(openingTag, 'applyAlignment')) === true ? readAlignmentStyle(xfXml) : undefined
  const protection =
    readBooleanAttribute(readAttribute(openingTag, 'applyProtection')) === true
      ? (readProtectionStyle(xfXml) ?? {})
      : readProtectionStyle(xfXml)
  return {
    fillId,
    fontId,
    fillApplied: isApplied(openingTag, 'applyFill', fillId),
    fontApplied: isApplied(openingTag, 'applyFont', fontId),
    ...(alignment ? { alignment } : {}),
    ...(protection !== undefined ? { protection } : {}),
  }
}

function readXfOpeningTag(xfXml: string): string | null {
  return /<(?:[A-Za-z_][\w.-]*:)?xf\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xfXml)?.[0] ?? null
}

function buildStylesByIndex(
  styleRefs: ReadonlyMap<number, StyleComponentRefs>,
  fills: ReadonlyMap<number, string> | null,
  fonts: ReadonlyMap<number, string> | null,
): Map<number, ImportedCellStyle> | null {
  if (!fills || !fonts) {
    return null
  }
  const styles = new Map<number, ImportedCellStyle>()
  for (const [styleIndex, refs] of styleRefs) {
    const fill = refs.fillApplied && refs.fillId !== null ? readFillStyle(fills.get(refs.fillId) ?? '') : undefined
    const font = refs.fontApplied && refs.fontId !== null ? readFontStyle(fonts.get(refs.fontId) ?? '') : undefined
    if (fill === null || font === null) {
      return null
    }
    const style: ImportedCellStyle = {
      ...(fill ? { fill } : {}),
      ...(font ? { font } : {}),
      ...(refs.alignment ? { alignment: refs.alignment } : {}),
      ...(refs.protection !== undefined ? { protection: refs.protection } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(styleIndex, style)
    }
  }
  return styles
}

function buildNumberFormatsByStyleIndex(
  styleFormatIds: ReadonlyMap<number, number>,
  customFormats: ReadonlyMap<number, string> | null,
): Map<number, string> | null {
  if (!customFormats) {
    return null
  }
  const formatsByStyleIndex = new Map<number, string>()
  for (const [styleIndex, formatId] of styleFormatIds) {
    const format = builtinNumberFormatCode(formatId) ?? customFormats.get(formatId)
    if (format) {
      formatsByStyleIndex.set(styleIndex, format)
    }
  }
  return formatsByStyleIndex
}

function readFillStyles(stylesXml: string): Array<ImportedCellStyle['fill'] | null> | null {
  const fillsXml = extractElementXml(stylesXml, 'fills')
  if (!fillsXml) {
    return []
  }
  return [...fillsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?fill\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?fill>/gu)].map((match) =>
    readFillStyle(match[0] ?? ''),
  )
}

function readFillStyle(fillXml: string): ImportedCellStyle['fill'] | null | undefined {
  const patternFill = extractElementXml(fillXml, 'patternFill')
  if (!patternFill) {
    return undefined
  }
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?patternFill\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(patternFill)?.[0]
  const patternType = openingTag ? readAttribute(openingTag, 'patternType') : undefined
  if (!patternType || patternType === 'none' || patternType === 'gray125') {
    return undefined
  }
  if (patternType !== 'solid') {
    return null
  }
  const color = readColor(patternFill, 'fgColor') ?? readColor(patternFill, 'bgColor')
  return color ? { backgroundColor: color } : undefined
}

function readFontStyles(stylesXml: string): Array<CellStyleFontSnapshot | null | undefined> | null {
  const fontsXml = extractElementXml(stylesXml, 'fonts')
  if (!fontsXml) {
    return []
  }
  return [...fontsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?font\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?font>/gu)].map((match) =>
    readFontStyle(match[0] ?? ''),
  )
}

function readFontStyle(fontXml: string): CellStyleFontSnapshot | null | undefined {
  const family = readElementValue(fontXml, 'name')
  const size = readElementNumberValue(fontXml, 'sz')
  const color = readColor(fontXml, 'color')
  const font: CellStyleFontSnapshot = {
    ...(family ? { family } : {}),
    ...(size ? { size } : {}),
    ...(hasBooleanElement(fontXml, 'b') ? { bold: true } : {}),
    ...(hasBooleanElement(fontXml, 'i') ? { italic: true } : {}),
    ...(hasBooleanElement(fontXml, 'u') ? { underline: true } : {}),
    ...(color ? { color } : {}),
  }
  return Object.keys(font).length > 0 ? font : undefined
}

function extractElementXml(xml: string, elementName: string): string | null {
  let pattern = elementTextCache.get(elementName)
  if (!pattern) {
    const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
    pattern = new RegExp(`<${qualifiedName}\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${qualifiedName}>)`, 'u')
    elementTextCache.set(elementName, pattern)
  }
  return pattern.exec(xml)?.[0] ?? null
}

function readElementValue(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return undefined
  }
  return readAttribute(elementXml, 'val')
}

function readElementNumberValue(xml: string, elementName: string): number | undefined {
  const value = readElementValue(xml, elementName)
  if (value === undefined) {
    return undefined
  }
  const number = Number(value)
  return Number.isFinite(number) && number > 0 ? number : undefined
}

function hasBooleanElement(xml: string, elementName: string): boolean {
  const elementXml = extractElementXml(xml, elementName)
  if (!elementXml) {
    return false
  }
  const value = readAttribute(elementXml, 'val')
  return value === undefined || value === '1' || value.toLocaleLowerCase('en-US') === 'true'
}

function readColor(xml: string, elementName: string): string | undefined {
  const elementXml = extractElementXml(xml, elementName)
  const rgb = elementXml ? readAttribute(elementXml, 'rgb') : undefined
  if (!rgb) {
    return undefined
  }
  const normalized = rgb.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{8}$/u.test(normalized)) {
    return `#${normalized.slice(2).toLocaleLowerCase('en-US')}`
  }
  if (/^[0-9a-fA-F]{6}$/u.test(normalized)) {
    return `#${normalized.toLocaleLowerCase('en-US')}`
  }
  return undefined
}

function readAlignmentStyle(xfXml: string): CellStyleAlignmentSnapshot | undefined {
  const alignmentXml = extractElementXml(xfXml, 'alignment')
  if (!alignmentXml) {
    return undefined
  }
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?alignment\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(alignmentXml)?.[0]
  if (!openingTag) {
    return undefined
  }
  const horizontal = readHorizontalAlignment(readAttribute(openingTag, 'horizontal'))
  const vertical = readVerticalAlignment(readAttribute(openingTag, 'vertical'))
  const indent = readNumberAttribute(openingTag, 'indent')
  const readingOrder = readNumberAttribute(openingTag, 'readingOrder')
  const textRotation = readNumberAttribute(openingTag, 'textRotation')
  const alignment: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(readBooleanAttribute(readAttribute(openingTag, 'wrapText')) === true ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
    ...(readBooleanAttribute(readAttribute(openingTag, 'shrinkToFit')) === true ? { shrinkToFit: true } : {}),
    ...(readingOrder !== null ? { readingOrder } : {}),
    ...(textRotation !== null ? { textRotation } : {}),
    ...(readBooleanAttribute(readAttribute(openingTag, 'justifyLastLine')) === true ? { justifyLastLine: true } : {}),
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined
}

function readProtectionStyle(xfXml: string): CellStyleProtectionSnapshot | undefined {
  const protectionXml = extractElementXml(xfXml, 'protection')
  if (!protectionXml) {
    return undefined
  }
  const openingTag = /<(?:[A-Za-z_][\w.-]*:)?protection\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(protectionXml)?.[0]
  if (!openingTag) {
    return undefined
  }
  const locked = readBooleanAttribute(readAttribute(openingTag, 'locked'))
  const hidden = readBooleanAttribute(readAttribute(openingTag, 'hidden'))
  return {
    ...(locked !== undefined ? { locked } : {}),
    ...(hidden !== undefined ? { hidden } : {}),
  }
}

function readHorizontalAlignment(value: string | undefined): CellHorizontalAlignment | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function readVerticalAlignment(value: string | undefined): CellVerticalAlignment | undefined {
  if (value === undefined) {
    return undefined
  }
  switch (value) {
    case 'top':
      return 'top'
    case 'center':
    case 'middle':
      return 'middle'
    case 'bottom':
      return 'bottom'
    case 'justify':
      return 'justify'
    case 'distributed':
      return 'distributed'
    default:
      return undefined
  }
}

function collectIndexedXmlElementsFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  parentName: string,
  childName: string,
  requiredIndexes: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions,
): Map<number, string> | null {
  if (requiredIndexes.size === 0) {
    return new Map()
  }
  const collector = new IndexedXmlElementCollector(parentName, childName, requiredIndexes, options)
  if (!readChunks((chunk) => collector.push(chunk))) {
    return null
  }
  return collector.finish()
}

function collectLargeSimpleWorkbookStyleComponentsFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  requiredFillIndexes: ReadonlySet<number>,
  requiredFontIndexes: ReadonlySet<number>,
  requiredCustomFormatIds: ReadonlySet<number>,
  options: LargeSimpleWorkbookStylesScanOptions,
): LargeSimpleWorkbookStyleComponents | null {
  const fillCollector = new IndexedXmlElementCollector('fills', 'fill', requiredFillIndexes, options)
  const fontCollector = new IndexedXmlElementCollector('fonts', 'font', requiredFontIndexes, options)
  const numberFormatCollector = new LargeSimpleNumberFormatCollector(requiredCustomFormatIds, options)
  if (
    !readChunks((chunk) => {
      fillCollector.push(chunk)
      fontCollector.push(chunk)
      numberFormatCollector.push(chunk)
    })
  ) {
    return null
  }
  return {
    fills: fillCollector.finish(),
    fonts: fontCollector.finish(),
    numberFormats: numberFormatCollector.finish(),
  }
}

class IndexedXmlElementCollector {
  private readonly decoder = new TextDecoder()
  private buffer = ''
  private index = 0
  private inParent = false
  private childIndex = 0
  private failed = false
  private skippingUnrequiredChild = false
  private readonly elements = new Map<number, string>()

  constructor(
    private readonly parentName: string,
    private readonly childName: string,
    private readonly requiredIndexes: ReadonlySet<number>,
    private readonly options: LargeSimpleWorkbookStylesScanOptions,
  ) {}

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0 || this.isComplete()) {
      return
    }
    this.buffer += this.decoder.decode(chunk, { stream: true })
    this.process(false)
    this.releaseBufferIfComplete()
    if (!this.isComplete()) {
      this.compact()
    }
    this.reportRetainedBufferLength()
  }

  finish(): Map<number, string> | null {
    if (this.failed) {
      return null
    }
    if (!this.isComplete()) {
      this.buffer += this.decoder.decode()
      this.process(true)
    }
    this.releaseBufferIfComplete()
    if (!this.isComplete()) {
      this.compact()
    }
    this.reportRetainedBufferLength()
    if (this.failed || this.elements.size !== this.requiredIndexes.size) {
      return null
    }
    return this.elements
  }

  private process(final: boolean): void {
    while (!this.failed && this.elements.size < this.requiredIndexes.size) {
      if (this.skippingUnrequiredChild) {
        if (!this.finishSkippingUnrequiredChild(final)) {
          return
        }
        continue
      }
      if (!this.inParent) {
        const parent = findNextOpeningTag(this.buffer, this.index, this.parentName)
        if (!parent) {
          this.index = Math.max(0, this.buffer.length - this.parentName.length - 4)
          return
        }
        const tagEnd = findStringTagEnd(this.buffer, parent.nameEnd)
        if (tagEnd === null) {
          if (final) {
            this.failed = true
          }
          this.index = parent.start
          return
        }
        if (isSelfClosingStringTag(this.buffer, tagEnd)) {
          this.index = tagEnd + 1
          continue
        }
        this.inParent = true
        this.childIndex = 0
        this.index = tagEnd + 1
        continue
      }
      const next = findNextParentBoundaryOrChild(this.buffer, this.index, this.parentName, this.childName)
      if (!next) {
        this.index = Math.max(0, this.buffer.length - Math.max(this.parentName.length, this.childName.length) - 4)
        return
      }
      if (next.kind === 'parent-close') {
        const tagEnd = findStringTagEnd(this.buffer, next.nameEnd)
        if (tagEnd === null) {
          if (final) {
            this.failed = true
          }
          this.index = next.start
          return
        }
        this.inParent = false
        this.index = tagEnd + 1
        continue
      }
      const tagEnd = findStringTagEnd(this.buffer, next.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = next.start
        return
      }
      const childStart = next.start
      const childEnd = isSelfClosingStringTag(this.buffer, tagEnd)
        ? tagEnd + 1
        : findClosingStringElementEnd(this.buffer, tagEnd + 1, this.childName)
      if (childEnd === null) {
        if (final) {
          this.failed = true
        }
        if (this.requiredIndexes.has(this.childIndex)) {
          this.index = childStart
        } else {
          this.skippingUnrequiredChild = true
          this.index = tagEnd + 1
        }
        return
      }
      if (this.requiredIndexes.has(this.childIndex)) {
        this.elements.set(this.childIndex, this.buffer.slice(childStart, childEnd))
      }
      this.childIndex += 1
      this.index = childEnd
    }
  }

  private isComplete(): boolean {
    return this.elements.size === this.requiredIndexes.size
  }

  private releaseBufferIfComplete(): void {
    if (!this.isComplete()) {
      return
    }
    this.buffer = ''
    this.index = 0
    this.skippingUnrequiredChild = false
  }

  private finishSkippingUnrequiredChild(final: boolean): boolean {
    const childEnd = findClosingStringElementEnd(this.buffer, this.index, this.childName)
    if (childEnd === null) {
      if (final) {
        this.failed = true
      }
      this.index = this.buffer.length
      return false
    }
    this.childIndex += 1
    this.index = childEnd
    this.skippingUnrequiredChild = false
    return true
  }

  private compact(): void {
    if (this.skippingUnrequiredChild) {
      const retainLength = closingStringTagRetainLength(this.childName)
      this.buffer = this.buffer.slice(Math.max(0, this.buffer.length - retainLength))
      this.index = 0
      return
    }
    if (this.index === 0) {
      return
    }
    if (this.index >= this.buffer.length) {
      this.buffer = ''
      this.index = 0
      return
    }
    this.buffer = this.buffer.slice(this.index)
    this.index = 0
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.length)
  }
}

class RequiredXfUnsupportedStyleScanner {
  private buffer: Uint8Array = emptyBytes
  private index = 0
  private inCellXfs = false
  private childIndex = 0
  private foundRequiredCount = 0
  private failed = false
  private unsupported = false
  private hasPotentialVisualStyles = false
  private tagNameEnd = 0

  constructor(
    private readonly requiredIndexes: ReadonlySet<number>,
    private readonly options: LargeSimpleWorkbookStylesScanOptions,
  ) {}

  push(chunk: Uint8Array): boolean {
    if (this.failed || this.unsupported || this.isComplete() || chunk.byteLength === 0) {
      return !this.failed && !this.unsupported && !this.isComplete()
    }
    this.append(chunk)
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
    return !this.failed && !this.unsupported && !this.isComplete()
  }

  finish(): LargeSimpleWorkbookStyleSupportScan | null {
    if (this.failed) {
      return null
    }
    if (!this.unsupported && !this.isComplete()) {
      this.process(true)
    }
    this.compact()
    this.reportRetainedBufferLength()
    if (this.failed) {
      return null
    }
    if (this.unsupported) {
      return { hasUnsupportedStyles: true, hasPotentialVisualStyles: this.hasPotentialVisualStyles }
    }
    return this.isComplete() ? { hasUnsupportedStyles: false, hasPotentialVisualStyles: this.hasPotentialVisualStyles } : null
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

  private process(final: boolean): void {
    while (!this.failed && !this.unsupported && !this.isComplete() && this.index < this.buffer.byteLength) {
      const tagStart = findByteInRange(this.buffer, this.index, this.buffer.byteLength, lessThan)
      if (tagStart === null) {
        this.index = this.buffer.byteLength
        return
      }
      if (!final && tagStart + 1 >= this.buffer.byteLength) {
        this.index = tagStart
        return
      }
      const closing = this.buffer[tagStart + 1] === slash
      const tagNameStart = tagStart + (closing ? 2 : 1)
      const tagCode = this.readTagCode(tagNameStart)
      if (tagCode === styleSupportTagUnknown) {
        if (!final && tagNameStart >= this.buffer.byteLength) {
          this.index = tagStart
          return
        }
        this.index = tagStart + 1
        continue
      }
      const tagEnd = findByteTagEnd(this.buffer, this.tagNameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = tagStart
        return
      }
      if (!this.inCellXfs) {
        if (!closing && tagCode === styleSupportTagCellXfs) {
          this.inCellXfs = !isSelfClosingByteTag(this.buffer, tagEnd)
        }
        this.index = tagEnd + 1
        continue
      }
      if (closing && tagCode === styleSupportTagCellXfs) {
        this.inCellXfs = false
        this.index = tagEnd + 1
        continue
      }
      if (!closing && tagCode === styleSupportTagXf) {
        if (this.requiredIndexes.has(this.childIndex)) {
          this.foundRequiredCount += 1
          const supportFlags = inspectRequiredXfOpeningTag(this.buffer, this.tagNameEnd, tagEnd)
          if ((supportFlags & xfSupportUnsupported) !== 0) {
            this.unsupported = true
            this.buffer = emptyBytes
            this.index = 0
            return
          }
          if ((supportFlags & xfSupportPotentialVisual) !== 0) {
            this.hasPotentialVisualStyles = true
          }
        }
        this.childIndex += 1
      }
      this.index = tagEnd + 1
    }
    if (this.isComplete() || this.unsupported) {
      this.buffer = emptyBytes
      this.index = 0
    }
  }

  private compact(): void {
    if (this.unsupported || this.isComplete()) {
      this.buffer = emptyBytes
      this.index = 0
      return
    }
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

  private isComplete(): boolean {
    return this.foundRequiredCount === this.requiredIndexes.size
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.byteLength)
  }

  private readTagCode(startIndex: number): number {
    const first = this.buffer[startIndex]
    if (first === undefined || first === 33 || first === slash || first === 63) {
      this.tagNameEnd = 0
      return styleSupportTagUnknown
    }
    let index = startIndex
    let localNameStart = startIndex
    while (index < this.buffer.byteLength && isXmlNameByte(this.buffer[index] ?? 0)) {
      if (this.buffer[index] === 58) {
        localNameStart = index + 1
      }
      index += 1
    }
    this.tagNameEnd = index
    if (index === localNameStart) {
      return styleSupportTagUnknown
    }
    if (asciiByteRangeEquals(this.buffer, localNameStart, index, 'cellXfs')) {
      return styleSupportTagCellXfs
    }
    return asciiByteRangeEquals(this.buffer, localNameStart, index, 'xf') ? styleSupportTagXf : styleSupportTagOther
  }
}
