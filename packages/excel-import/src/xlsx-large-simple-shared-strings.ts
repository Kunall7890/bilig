import type { WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'
import { decodeXmlText } from './xlsx-large-simple-xml-text.js'
import type { ImportedWorkbookArenaDedupeMode } from './xlsx-large-simple-arena-types.js'
import type { LargeSimpleSharedStringIndexSet } from './xlsx-large-simple-shared-string-indexes.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'

export interface LargeSimpleSharedStringEntry {
  readonly text: string
  readonly xml?: string
  readonly rich: boolean
}

export interface LargeSimpleSharedStringTable {
  readonly length: number
  readonly [index: number]: LargeSimpleSharedStringEntry | undefined
}

export type LargeSimpleSharedStrings = readonly LargeSimpleSharedStringEntry[] | LargeSimpleSharedStringTable

const sharedStringElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?si\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u
const partialSharedStringTagRetainLength = 256
const sparseReferencedSharedStringDensityDivisor = 64
const lessThanByte = 60
const slashByte = 47
const greaterThanByte = 62
const doubleQuoteByte = 34
const singleQuoteByte = 39
const richSharedStringTagUnknown = 0
const richSharedStringTagSi = 1
const richSharedStringTagRun = 2
const richSharedStringTagOther = 3
const emptyBytes = new Uint8Array(0)

export interface LargeSimpleReferencedSharedStringScanOptions {
  readonly onRetainedBufferLength?: (length: number) => void
  readonly stringPool?: ImportedWorkbookStringPool
  readonly deduplicateText?: ImportedWorkbookArenaDedupeMode
  readonly dedupeMaxEntries?: number
}

type LargeSimpleChunkConsumer = (chunk: Uint8Array) => boolean | void

interface LargeSimpleSharedStringReadOptions extends LargeSimpleReferencedSharedStringScanOptions {
  readonly plainEntryPool?: LargeSimpleSharedStringEntryPool
}

interface LargeSimpleSharedStringEntryPool {
  readonly entries: Map<string, LargeSimpleSharedStringEntry>
  readonly keys: string[]
  evictionIndex: number
}

export function readLargeSimpleSharedStrings(
  sharedStringsXml: string,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStringEntry[] {
  const readOptions = sharedStringReadOptions(options)
  return [...sharedStringsXml.matchAll(sharedStringElementPattern)].map((match) => {
    const xml = match[0]
    return readLargeSimpleSharedStringEntry(xml, readOptions)
  })
}

export function readLargeSimpleReferencedSharedStringsFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  referencedIndexes: LargeSimpleSharedStringIndexSet,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStrings | null {
  if (referencedIndexes.size === 0) {
    return []
  }
  const scanner = new LargeSimpleSharedStringChunkScanner(referencedIndexes, options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export function readAllLargeSimpleSharedStringsFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  options: LargeSimpleReferencedSharedStringScanOptions = {},
): LargeSimpleSharedStrings | null {
  const scanner = new LargeSimpleAllSharedStringChunkScanner(options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export function hasAnyLargeSimpleRichSharedStringFromChunks(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => boolean,
  options: Pick<LargeSimpleReferencedSharedStringScanOptions, 'onRetainedBufferLength'> = {},
): boolean | null {
  const scanner = new LargeSimpleRichSharedStringPresenceScanner(options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

export async function hasAnyLargeSimpleRichSharedStringFromChunksAsync(
  readChunks: (onChunk: LargeSimpleChunkConsumer) => Promise<boolean>,
  options: Pick<LargeSimpleReferencedSharedStringScanOptions, 'onRetainedBufferLength'> = {},
): Promise<boolean | null> {
  const scanner = new LargeSimpleRichSharedStringPresenceScanner(options)
  if (!(await readChunks((chunk) => scanner.push(chunk)))) {
    return null
  }
  return scanner.finish()
}

export function createLargeSimpleSharedStringSubset(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: LargeSimpleSharedStringIndexSet,
): LargeSimpleSharedStrings | null {
  if (referencedIndexes.size === 0) {
    return []
  }
  const entries = new Map<number, LargeSimpleSharedStringEntry>()
  let maxReferencedIndex = 0
  for (const index of referencedIndexes) {
    const entry = sharedStrings[index]
    if (!entry) {
      return null
    }
    entries.set(index, entry)
    maxReferencedIndex = Math.max(maxReferencedIndex, index)
  }
  return createReferencedSharedStringTable(entries, maxReferencedIndex, { preferSparse: true })
}

export function hasReferencedLargeSimpleRichSharedStrings(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: LargeSimpleSharedStringIndexSet,
): boolean {
  for (const index of referencedIndexes) {
    if (sharedStrings[index]?.rich) {
      return true
    }
  }
  return false
}

export function collectReferencedLargeSimpleRichSharedStringIndexes(
  sharedStrings: LargeSimpleSharedStrings,
  referencedIndexes: LargeSimpleSharedStringIndexSet,
): Set<number> | null {
  const richIndexes = new Set<number>()
  for (const index of referencedIndexes) {
    const entry = sharedStrings[index]
    if (!entry) {
      return null
    }
    if (entry.rich) {
      richIndexes.add(index)
    }
  }
  return richIndexes
}

export function readLargeSimpleRichTextCellArtifact(
  address: string,
  openingTag: string,
  cellXml: string,
  sharedStrings: LargeSimpleSharedStrings,
): WorkbookRichTextCellSnapshot | undefined {
  const type = readXmlAttribute(openingTag, 't')
  if (type === 's') {
    const entry = sharedStrings[readSharedStringIndex(cellXml) ?? -1]
    return entry?.rich
      ? {
          address,
          text: entry.text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        }
      : undefined
  }
  if (type === 'inlineStr') {
    const inlineStringXml = readStringElement(cellXml, 'is')
    if (inlineStringXml && richTextRunPattern.test(inlineStringXml)) {
      return {
        address,
        text: stringItemText(inlineStringXml),
        storage: 'inlineString',
        xml: inlineStringXml,
      }
    }
  }
  return undefined
}

class LargeSimpleSharedStringChunkScanner {
  private readonly decoder = new TextDecoder()
  private readonly denseEntries: LargeSimpleSharedStringEntry[] | null
  private readonly sparseEntries: Map<number, LargeSimpleSharedStringEntry> | null
  private buffer = ''
  private index = 0
  private sharedStringIndex = 0
  private foundReferencedCount = 0
  private failed = false
  private skippingUnreferencedElementName: string | null = null
  private readonly maxReferencedIndex: number
  private readonly readOptions: LargeSimpleSharedStringReadOptions

  constructor(
    private readonly referencedIndexes: LargeSimpleSharedStringIndexSet,
    private readonly options: LargeSimpleReferencedSharedStringScanOptions,
  ) {
    this.readOptions = sharedStringReadOptions(options)
    let maxIndex = 0
    for (const index of referencedIndexes) {
      maxIndex = Math.max(maxIndex, index)
    }
    this.maxReferencedIndex = maxIndex
    if (maxIndex + 1 <= referencedIndexes.size * sparseReferencedSharedStringDensityDivisor) {
      this.denseEntries = []
      this.denseEntries.length = maxIndex + 1
      this.sparseEntries = null
    } else {
      this.denseEntries = null
      this.sparseEntries = new Map()
    }
  }

  push(chunk: Uint8Array): boolean {
    if (this.failed || this.isComplete() || chunk.byteLength === 0) {
      return !this.failed && !this.isComplete()
    }
    this.buffer += this.decoder.decode(chunk, { stream: true })
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
    return !this.failed && !this.isComplete()
  }

  finish(): LargeSimpleSharedStrings | null {
    if (this.failed) {
      return null
    }
    if (!this.isComplete()) {
      this.buffer += this.decoder.decode()
      this.process(true)
      this.compact()
      this.reportRetainedBufferLength()
    }
    if (this.failed) {
      return null
    }
    for (const index of this.referencedIndexes) {
      if (!this.hasEntry(index)) {
        return null
      }
    }
    return this.denseEntries ?? createReferencedSharedStringTable(this.sparseEntries ?? new Map(), this.maxReferencedIndex)
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.length) {
      if (this.skippingUnreferencedElementName !== null) {
        if (!this.finishSkippingUnreferencedElement(final)) {
          return
        }
        continue
      }
      const opening = findNextElementOpening(this.buffer, 'si', this.index)
      if (!opening) {
        this.index = final ? this.buffer.length : Math.max(this.index, this.buffer.length - partialSharedStringTagRetainLength)
        return
      }
      const tagEnd = findTagEnd(this.buffer, opening.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return
      }
      const xmlEnd = isSelfClosingTag(this.buffer, tagEnd) ? tagEnd + 1 : findClosingElementEnd(this.buffer, opening.name, tagEnd + 1)
      if (xmlEnd === null) {
        if (final) {
          this.failed = true
        }
        if (!this.referencedIndexes.has(this.sharedStringIndex)) {
          this.skippingUnreferencedElementName = opening.name
          this.index = tagEnd + 1
        }
        return
      }
      if (this.referencedIndexes.has(this.sharedStringIndex)) {
        this.setEntry(
          this.sharedStringIndex,
          readLargeSimpleSharedStringEntryFromRange(this.buffer, opening.start, xmlEnd, this.readOptions),
        )
      }
      this.sharedStringIndex += 1
      this.index = xmlEnd
    }
  }

  private finishSkippingUnreferencedElement(final: boolean): boolean {
    const elementName = this.skippingUnreferencedElementName
    if (elementName === null) {
      return true
    }
    const xmlEnd = findClosingElementEnd(this.buffer, elementName, this.index)
    if (xmlEnd === null) {
      if (final) {
        this.failed = true
      }
      this.index = this.buffer.length
      return false
    }
    this.sharedStringIndex += 1
    this.index = xmlEnd
    this.skippingUnreferencedElementName = null
    return true
  }

  private compact(): void {
    if (this.skippingUnreferencedElementName !== null) {
      const retainLength = closingTagRetainLength(this.skippingUnreferencedElementName)
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

  private hasEntry(index: number): boolean {
    return this.denseEntries ? this.denseEntries[index] !== undefined : (this.sparseEntries?.has(index) ?? false)
  }

  private setEntry(index: number, entry: LargeSimpleSharedStringEntry): void {
    if (!this.hasEntry(index)) {
      this.foundReferencedCount += 1
    }
    if (this.denseEntries) {
      this.denseEntries[index] = entry
      return
    }
    this.sparseEntries?.set(index, entry)
  }

  private isComplete(): boolean {
    return this.foundReferencedCount === this.referencedIndexes.size
  }
}

class LargeSimpleRichSharedStringPresenceScanner {
  private buffer: Uint8Array = emptyBytes
  private index = 0
  private inSharedString = false
  private failed = false
  private found = false
  private tagNameEnd = 0

  constructor(private readonly options: Pick<LargeSimpleReferencedSharedStringScanOptions, 'onRetainedBufferLength'>) {}

  push(chunk: Uint8Array): boolean {
    if (this.failed || this.found || chunk.byteLength === 0) {
      return !this.failed && !this.found
    }
    this.append(chunk)
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
    return !this.failed && !this.found
  }

  finish(): boolean | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    this.compact()
    this.reportRetainedBufferLength()
    return this.failed ? null : this.found
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
    while (!this.failed && !this.found && this.index < this.buffer.byteLength) {
      const tagStart = findByteInRange(this.buffer, this.index, this.buffer.byteLength, lessThanByte)
      if (tagStart === null) {
        this.index = this.buffer.byteLength
        return
      }
      if (!final && tagStart + 1 >= this.buffer.byteLength) {
        this.index = tagStart
        return
      }
      const closing = this.buffer[tagStart + 1] === slashByte
      const tagNameStart = tagStart + (closing ? 2 : 1)
      const tagCode = this.readElementNameCode(tagNameStart)
      if (tagCode === richSharedStringTagUnknown) {
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
      if (!this.inSharedString) {
        if (!closing && tagCode === richSharedStringTagSi) {
          this.inSharedString = !isSelfClosingByteTag(this.buffer, tagEnd)
        }
        this.index = tagEnd + 1
        continue
      }
      if (closing && tagCode === richSharedStringTagSi) {
        this.inSharedString = false
        this.index = tagEnd + 1
        continue
      }
      if (!closing && tagCode === richSharedStringTagRun) {
        this.found = true
        return
      }
      this.index = tagEnd + 1
    }
    if (final && this.inSharedString) {
      this.failed = true
    }
  }

  private compact(): void {
    if (this.found) {
      this.buffer = emptyBytes
      this.index = 0
      this.inSharedString = false
      return
    }
    if (this.index === 0) {
      return
    }
    const retainLength = this.inSharedString ? partialSharedStringTagRetainLength : 0
    if (this.index >= this.buffer.byteLength) {
      this.buffer = retainLength > 0 ? this.buffer.subarray(Math.max(0, this.buffer.byteLength - retainLength)) : emptyBytes
      this.index = 0
      return
    }
    const startIndex = retainLength > 0 ? Math.max(0, this.index - retainLength) : this.index
    this.buffer = this.buffer.subarray(startIndex)
    this.index -= startIndex
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.byteLength)
  }

  private readElementNameCode(startIndex: number): number {
    const first = this.buffer[startIndex]
    if (first === undefined || first === 33 || first === slashByte || first === 63) {
      this.tagNameEnd = 0
      return richSharedStringTagUnknown
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
      return richSharedStringTagUnknown
    }
    if (localNameMatches(this.buffer, localNameStart, index, 'si')) {
      return richSharedStringTagSi
    }
    return localNameMatches(this.buffer, localNameStart, index, 'r') ? richSharedStringTagRun : richSharedStringTagOther
  }
}

function findByteTagEnd(bytes: Uint8Array, startIndex: number): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < bytes.byteLength; index += 1) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuoteByte || byte === singleQuoteByte) {
      quote = byte
      continue
    }
    if (byte === greaterThanByte) {
      return index
    }
  }
  return null
}

function isSelfClosingByteTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slashByte
}

function findByteInRange(bytes: Uint8Array, startIndex: number, endIndex: number, target: number): number | null {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === target) {
      return index
    }
  }
  return null
}

function isXmlNameByte(byte: number): boolean {
  return (
    (byte >= 65 && byte <= 90) ||
    (byte >= 97 && byte <= 122) ||
    (byte >= 48 && byte <= 57) ||
    byte === 45 ||
    byte === 46 ||
    byte === 58 ||
    byte === 95
  )
}

function localNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, localName: string): boolean {
  if (endIndex - startIndex !== localName.length) {
    return false
  }
  for (let index = 0; index < localName.length; index += 1) {
    if (bytes[startIndex + index] !== localName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

class LargeSimpleAllSharedStringChunkScanner {
  private readonly decoder = new TextDecoder()
  private readonly entries: LargeSimpleSharedStringEntry[] = []
  private buffer = ''
  private index = 0
  private failed = false
  private readonly readOptions: LargeSimpleSharedStringReadOptions

  constructor(private readonly options: LargeSimpleReferencedSharedStringScanOptions) {
    this.readOptions = sharedStringReadOptions(options)
  }

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.buffer += this.decoder.decode(chunk, { stream: true })
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
  }

  finish(): LargeSimpleSharedStrings | null {
    if (this.failed) {
      return null
    }
    this.buffer += this.decoder.decode()
    this.process(true)
    this.compact()
    this.reportRetainedBufferLength()
    return this.failed ? null : this.entries
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.length) {
      const opening = findNextElementOpening(this.buffer, 'si', this.index)
      if (!opening) {
        this.index = final ? this.buffer.length : Math.max(this.index, this.buffer.length - partialSharedStringTagRetainLength)
        return
      }
      const tagEnd = findTagEnd(this.buffer, opening.nameEnd)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = opening.start
        return
      }
      const xmlEnd = isSelfClosingTag(this.buffer, tagEnd) ? tagEnd + 1 : findClosingElementEnd(this.buffer, opening.name, tagEnd + 1)
      if (xmlEnd === null) {
        if (final) {
          this.failed = true
        }
        this.index = opening.start
        return
      }
      this.entries.push(readLargeSimpleSharedStringEntryFromRange(this.buffer, opening.start, xmlEnd, this.readOptions))
      this.index = xmlEnd
    }
  }

  private compact(): void {
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

function createReferencedSharedStringTable(
  entries: ReadonlyMap<number, LargeSimpleSharedStringEntry>,
  maxReferencedIndex: number,
  options: { readonly preferSparse?: boolean } = {},
): LargeSimpleSharedStrings {
  const length = maxReferencedIndex + 1
  if (options.preferSparse !== true && length <= entries.size * sparseReferencedSharedStringDensityDivisor) {
    const output: LargeSimpleSharedStringEntry[] = []
    output.length = length
    for (const [index, entry] of entries) {
      output[index] = entry
    }
    return output
  }
  return new Proxy(
    { length },
    {
      get(target, property) {
        if (property === 'length') {
          return target.length
        }
        return typeof property === 'string' && isArrayIndexProperty(property) ? entries.get(Number(property)) : undefined
      },
      has(_target, property) {
        return property === 'length' || (typeof property === 'string' && isArrayIndexProperty(property) && entries.has(Number(property)))
      },
    },
  ) as LargeSimpleSharedStringTable
}

function isArrayIndexProperty(property: string): boolean {
  if (property.length === 0 || !/^(?:0|[1-9][0-9]*)$/u.test(property)) {
    return false
  }
  const index = Number(property)
  return Number.isSafeInteger(index) && index >= 0 && index < 2 ** 32 - 1
}

function closingTagRetainLength(elementName: string): number {
  return Math.max(partialSharedStringTagRetainLength, elementName.length + 4)
}

function readLargeSimpleSharedStringEntry(xml: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  const rich = richTextRunPattern.test(xml)
  if (!rich) {
    return internPlainSharedStringEntry(internSharedStringText(stringItemText(xml), options), options)
  }
  return lazyRichSharedStringEntry(xml, options)
}

function readLargeSimpleSharedStringEntryFromRange(
  xml: string,
  startIndex: number,
  endIndex: number,
  options: LargeSimpleSharedStringReadOptions,
): LargeSimpleSharedStringEntry {
  if (hasRichTextRunInRange(xml, startIndex, endIndex)) {
    return lazyRichSharedStringEntry(xml.slice(startIndex, endIndex), options)
  }
  return internPlainSharedStringEntry(internSharedStringText(stringItemTextInRange(xml, startIndex, endIndex), options), options)
}

function lazyRichSharedStringEntry(xml: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  let text: string | undefined
  return {
    rich: true,
    xml,
    get text() {
      text ??= internSharedStringText(stringItemText(xml), options)
      return text
    },
  }
}

function internSharedStringText(value: string, options: LargeSimpleReferencedSharedStringScanOptions): string {
  const mode = options.deduplicateText ?? 'bounded'
  if (mode === false || !options.stringPool) {
    return value
  }
  if (mode === 'bounded') {
    return options.stringPool.internBounded(value, options.dedupeMaxEntries ?? 8192)
  }
  return options.stringPool.intern(value)
}

function sharedStringReadOptions(options: LargeSimpleReferencedSharedStringScanOptions): LargeSimpleSharedStringReadOptions {
  if ((options.deduplicateText ?? 'bounded') === false) {
    return options
  }
  return {
    ...options,
    plainEntryPool: {
      entries: new Map(),
      keys: [],
      evictionIndex: 0,
    },
  }
}

function internPlainSharedStringEntry(text: string, options: LargeSimpleSharedStringReadOptions): LargeSimpleSharedStringEntry {
  const pool = options.plainEntryPool
  if (!pool) {
    return { text, rich: false }
  }
  const existing = pool.entries.get(text)
  if (existing) {
    return existing
  }
  const entry = { text, rich: false } satisfies LargeSimpleSharedStringEntry
  pool.entries.set(text, entry)
  pool.keys.push(text)
  if ((options.deduplicateText ?? 'bounded') === 'bounded') {
    evictPlainSharedStringEntries(pool, options.dedupeMaxEntries ?? 8192)
  }
  return entry
}

function evictPlainSharedStringEntries(pool: LargeSimpleSharedStringEntryPool, maxEntries: number): void {
  const limit = Math.max(0, Math.trunc(maxEntries))
  while (pool.keys.length - pool.evictionIndex > limit) {
    const key = pool.keys[pool.evictionIndex]
    pool.evictionIndex += 1
    if (key !== undefined) {
      pool.entries.delete(key)
    }
  }
  if (pool.evictionIndex > limit && pool.evictionIndex * 2 > pool.keys.length) {
    pool.keys.splice(0, pool.evictionIndex)
    pool.evictionIndex = 0
  }
}

function readSharedStringIndex(cellXml: string): number | null {
  const rawValue = readElementText(cellXml, 'v')?.trim()
  if (!rawValue) {
    return null
  }
  const index = Number(decodeXmlText(rawValue))
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function readStringElement(xml: string, elementName: 'is'): string | null {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)?.[0] ?? null
}

function readElementText(xml: string, elementName: 'v'): string | null {
  return (
    new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>([\\s\\S]*?)</(?:[A-Za-z_][\\w.-]*:)?${elementName}>`, 'u').exec(xml)?.[1] ??
    null
  )
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function findNextElementOpening(
  xml: string,
  localName: string,
  startIndex: number,
): { readonly start: number; readonly name: string; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < xml.length) {
    const openingStart = xml.indexOf('<', index)
    if (openingStart === -1) {
      return null
    }
    const name = readElementName(xml, openingStart + 1)
    if (!name) {
      index = openingStart + 1
      continue
    }
    if (readLocalName(name.name) === localName) {
      return {
        start: openingStart,
        name: name.name,
        nameEnd: name.end,
      }
    }
    index = name.end
  }
  return null
}

function readElementName(xml: string, startIndex: number): { readonly name: string; readonly end: number } | null {
  const first = xml[startIndex]
  if (!first || first === '/' || first === '?' || first === '!') {
    return null
  }
  let index = startIndex
  while (index < xml.length && isXmlNameCharacter(xml.charCodeAt(index))) {
    index += 1
  }
  return index > startIndex ? { name: xml.slice(startIndex, index), end: index } : null
}

function readLocalName(name: string): string {
  return name.includes(':') ? name.slice(name.lastIndexOf(':') + 1) : name
}

function findTagEnd(xml: string, startIndex: number): number | null {
  let quote: string | null = null
  for (let index = startIndex; index < xml.length; index += 1) {
    const character = xml[index]
    if (quote) {
      if (character === quote) {
        quote = null
      }
      continue
    }
    if (character === '"' || character === "'") {
      quote = character
      continue
    }
    if (character === '>') {
      return index
    }
  }
  return null
}

function isSelfClosingTag(xml: string, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(xml.charCodeAt(index))) {
    index -= 1
  }
  return xml[index] === '/'
}

function findClosingElementEnd(xml: string, name: string, startIndex: number): number | null {
  const closingStart = xml.indexOf(`</${name}`, startIndex)
  if (closingStart === -1) {
    return null
  }
  const closingNameEnd = closingStart + name.length + 2
  const next = xml[closingNameEnd]
  if (next !== '>' && !isAsciiWhitespace(next?.charCodeAt(0) ?? 0)) {
    return findClosingElementEnd(xml, name, closingNameEnd)
  }
  const tagEnd = findTagEnd(xml, closingNameEnd)
  return tagEnd === null ? null : tagEnd + 1
}

function stringItemText(xml: string): string {
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
    .map((match) => decodeExcelEscapedText(decodeXmlText(match[1] ?? '')))
    .join('')
}

function stringItemTextInRange(xml: string, startIndex: number, endIndex: number): string {
  const parts: string[] = []
  let index = startIndex
  while (index < endIndex) {
    const opening = findNextElementOpening(xml, 't', index)
    if (!opening || opening.start >= endIndex) {
      break
    }
    const tagEnd = findTagEnd(xml, opening.nameEnd)
    if (tagEnd === null || tagEnd >= endIndex) {
      break
    }
    if (isSelfClosingTag(xml, tagEnd)) {
      index = tagEnd + 1
      continue
    }
    const closingStart = findClosingElementStart(xml, opening.name, tagEnd + 1, endIndex)
    if (closingStart === null) {
      break
    }
    parts.push(decodeExcelEscapedText(decodeXmlText(xml.slice(tagEnd + 1, closingStart))))
    const closingEnd = findTagEnd(xml, closingStart + opening.name.length + 2)
    index = closingEnd === null ? endIndex : closingEnd + 1
  }
  return parts.join('')
}

function hasRichTextRunInRange(xml: string, startIndex: number, endIndex: number): boolean {
  let index = startIndex
  while (index < endIndex) {
    const openingStart = xml.indexOf('<', index)
    if (openingStart === -1 || openingStart >= endIndex) {
      return false
    }
    const name = readElementName(xml, openingStart + 1)
    if (!name) {
      index = openingStart + 1
      continue
    }
    if (name.end <= endIndex && readLocalName(name.name) === 'r') {
      return true
    }
    index = name.end
  }
  return false
}

function findClosingElementStart(xml: string, name: string, startIndex: number, endIndex: number): number | null {
  const closingPrefix = `</${name}`
  let searchIndex = startIndex
  while (searchIndex < endIndex) {
    const closingStart = xml.indexOf(closingPrefix, searchIndex)
    if (closingStart === -1 || closingStart >= endIndex) {
      return null
    }
    const closingNameEnd = closingStart + name.length + 2
    const next = xml[closingNameEnd]
    if (closingNameEnd < endIndex && (next === '>' || isAsciiWhitespace(next?.charCodeAt(0) ?? 0))) {
      return closingStart
    }
    searchIndex = closingNameEnd
  }
  return null
}

function isAsciiWhitespace(value: number): boolean {
  return value === 9 || value === 10 || value === 12 || value === 13 || value === 32
}

function isXmlNameCharacter(value: number): boolean {
  return (
    (value >= 65 && value <= 90) ||
    (value >= 97 && value <= 122) ||
    (value >= 48 && value <= 57) ||
    value === 45 ||
    value === 46 ||
    value === 58 ||
    value === 95
  )
}
