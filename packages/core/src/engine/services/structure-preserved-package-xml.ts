import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  renameFormulaSheetReferences,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import type { WorkbookPreservedMetadataRecord } from '../../workbook-metadata-types.js'

const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu
const connectionElementPattern = /<((?:[A-Za-z_][\w.-]*:)?connection)\b([^>]*?)(?:\/>|>[\s\S]*?<\/\1>)/gu
const pivotTableDefinitionLocationPattern = /<((?:[A-Za-z_][\w.-]*:)?location)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu
const worksheetSourcePattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu
const worksheetSourceElementPattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b[^>]*\/?>/gu

export function relationshipPartPathForPackagePart(partPath: string): string {
  const normalizedPath = normalizePivotPackagePath(partPath)
  const slashIndex = normalizedPath.lastIndexOf('/')
  return slashIndex >= 0
    ? `${normalizedPath.slice(0, slashIndex)}/_rels/${normalizedPath.slice(slashIndex + 1)}.rels`
    : `_rels/${normalizedPath}.rels`
}

export function packageRelationshipTargetPaths(basePartPath: string, relationshipsXml: string): Set<string> {
  const output = new Set<string>()
  for (const match of relationshipsXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?Relationship)\b([^>]*?)\/?>/gu)) {
    const attributes = match[2] ?? ''
    const target = readXmlAttribute(attributes, 'Target')
    const targetMode = readXmlAttribute(attributes, 'TargetMode')
    if (target && targetMode !== 'External') {
      output.add(normalizePivotPackagePath(resolvePivotRelationshipTarget(basePartPath, target)))
    }
  }
  return output
}

export function resolvePivotRelationshipTarget(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const segments = basePartPath.split('/')
  segments.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      segments.pop()
    } else if (segment !== '.' && segment.length > 0) {
      segments.push(segment)
    }
  }
  return segments.join('/')
}

export function normalizePivotPackagePath(path: string): string {
  return path.replace(/^\/+/u, '').replace(/\\/gu, '/')
}

export function normalizePackagePath(path: string): string {
  return normalizePivotPackagePath(path)
}

export function readXmlAttribute(attributes: string, name: string): string | undefined {
  return new RegExp(`\\b${name}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2]
}

export function readXmlAttributeMatch(
  attributes: string,
  name: string,
): { readonly raw: string; readonly quote: string; readonly value: string } | undefined {
  const match = new RegExp(`\\b${name}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)
  if (!match || !match[0] || !match[1]) {
    return undefined
  }
  return { raw: match[0], quote: match[1], value: match[2] ?? '' }
}

export function unescapeXmlAttribute(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}

export function escapeXmlAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll("'", '&apos;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

export function preservedPackagePartText(part: { readonly dataBase64: string; readonly byteLength: number }): string | undefined {
  const bytes = decodeBase64(part.dataBase64)
  return bytes.byteLength === part.byteLength ? new TextDecoder().decode(bytes) : undefined
}

export function rewritePreservedTextPackagePart<
  T extends { readonly storage: 'base64'; readonly dataBase64: string; readonly byteLength: number },
>(part: T, text: string): T {
  const bytes = new TextEncoder().encode(text)
  return {
    ...part,
    dataBase64: encodeBase64(bytes),
    byteLength: bytes.byteLength,
  }
}

export function renameSheetName(value: string, oldSheetName: string, newSheetName: string): string {
  return value === oldSheetName ? newSheetName : value
}

export function renameFormulaText(formula: string, oldSheetName: string, newSheetName: string): string {
  try {
    return renameFormulaSheetReferences(formula, oldSheetName, newSheetName)
  } catch {
    return formula
  }
}

export function rewriteCellReferenceForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  if (!isCellReference(address)) {
    return address
  }
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  return rewritten ? clipAddressToSheetGrid(rewritten) : undefined
}

export function cellReferenceTouchesAxisDelete(address: string, axis: 'row' | 'column', start: number): boolean {
  if (!isCellReference(address)) {
    return false
  }
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return false
  }
  return (axis === 'row' ? parsed[0] : parsed[1]) >= start
}

export function isCellReference(address: string): boolean {
  return cellReferencePattern.test(address)
}

export function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

export function hasConnectionXmlEntries(xml: string): boolean {
  connectionElementPattern.lastIndex = 0
  const found = connectionElementPattern.test(xml)
  connectionElementPattern.lastIndex = 0
  return found
}

export function removeConnectionXmlEntries(xml: string, connectionIds: ReadonlySet<string>): string {
  connectionElementPattern.lastIndex = 0
  const nextXml = xml.replace(connectionElementPattern, (source: string, _tagName: string, attributes: string) => {
    const id = readXmlAttribute(attributes, 'id')
    return id && connectionIds.has(id) ? '' : source
  })
  return rewriteConnectionsCount(nextXml)
}

export function rewriteConnectionsCount(xml: string): string {
  connectionElementPattern.lastIndex = 0
  const count = [...xml.matchAll(connectionElementPattern)].length
  return xml.replace(/<((?:[A-Za-z_][\w.-]*:)?connections)\b([^>]*?)>/u, (source: string, _tagName: string, attributes: string) => {
    const countAttribute = readXmlAttributeMatch(attributes, 'count')
    if (!countAttribute) {
      return source
    }
    return source.replace(countAttribute.raw, `count=${countAttribute.quote}${String(count)}${countAttribute.quote}`)
  })
}

export function removeWorkbookPivotCacheEntries(xml: string | undefined, cacheIds: ReadonlySet<string>): string | undefined {
  if (!xml || cacheIds.size === 0) {
    return xml
  }
  const nextXml = xml.replace(/<((?:[A-Za-z_][\w.-]*:)?pivotCache)\b([^>]*?)\/?>/gu, (source: string, _tag: string, attributes: string) => {
    const cacheId = readXmlAttribute(attributes, 'cacheId')
    return cacheId && cacheIds.has(cacheId) ? '' : source
  })
  if (nextXml === xml) {
    return xml
  }
  return /<((?:[A-Za-z_][\w.-]*:)?pivotCache)\b/u.test(nextXml) ? nextXml : undefined
}

export function hasPivotArtifacts(artifacts: NonNullable<WorkbookPreservedMetadataRecord['pivotArtifacts']>): boolean {
  return artifacts.parts.length > 0 || Boolean(artifacts.workbookPivotCachesXml) || (artifacts.workbookRelationships?.length ?? 0) > 0
}

export function rewritePivotTableDefinitionLocationRefsForStructuralTransform(
  xml: string,
  transform: StructuralAxisTransform,
): string | undefined {
  let failed = false
  const nextXml = xml.replace(
    pivotTableDefinitionLocationPattern,
    (source: string, _tagName: string, _attributes: string, quote: string, ref: string) => {
      const rewrittenRef = rewritePivotLocationRefForStructuralTransform(ref, transform)
      if (!rewrittenRef) {
        failed = true
        return source
      }
      return source.replace(`ref=${quote}${ref}${quote}`, `ref=${quote}${rewrittenRef}${quote}`)
    },
  )
  return failed ? undefined : nextXml
}

export function rewritePivotCacheWorksheetSourceRefsForStructuralTransform(
  xml: string,
  sheetName: string,
  transform: StructuralAxisTransform,
): string | undefined {
  let failed = false
  const nextXml = xml.replace(
    worksheetSourcePattern,
    (source: string, _tagName: string, attributes: string, quote: string, ref: string) => {
      const sourceSheetName = readXmlAttribute(attributes, 'sheet')
      if (sourceSheetName === undefined || unescapeXmlAttribute(sourceSheetName) !== sheetName) {
        return source
      }
      const rewrittenRef = rewritePivotLocationRefForStructuralTransform(ref, transform)
      if (!rewrittenRef) {
        failed = true
        return source
      }
      return source.replace(`ref=${quote}${ref}${quote}`, `ref=${quote}${rewrittenRef}${quote}`)
    },
  )
  return failed ? undefined : nextXml
}

export function renamePivotCacheWorksheetSourceSheetReferences(xml: string, oldSheetName: string, newSheetName: string): string {
  return xml.replace(worksheetSourceElementPattern, (source: string) => {
    const attribute = readXmlAttributeMatch(source, 'sheet')
    if (!attribute || unescapeXmlAttribute(attribute.value) !== oldSheetName) {
      return source
    }
    return source.replace(attribute.raw, `sheet=${attribute.quote}${escapeXmlAttribute(newSheetName)}${attribute.quote}`)
  })
}

function rewritePivotLocationRefForStructuralTransform(ref: string, transform: StructuralAxisTransform): string | undefined {
  const [startAddress, endAddress = startAddress] = ref.split(':')
  if (!startAddress || !endAddress || !isCellReference(startAddress) || !isCellReference(endAddress)) {
    return ref
  }
  const rewritten = rewriteRangeForStructuralTransform(startAddress, endAddress, transform)
  if (!rewritten) {
    return undefined
  }
  const start = clipAddressToSheetGrid(rewritten.startAddress)
  const end = clipAddressToSheetGrid(rewritten.endAddress)
  if (!start || !end) {
    return undefined
  }
  return start === end ? start : `${start}:${end}`
}

function encodeBase64(bytes: Uint8Array): string {
  return globalThis.btoa(encodeBinaryString(bytes))
}

function decodeBase64(dataBase64: string): Uint8Array {
  return decodeBinaryString(globalThis.atob(dataBase64))
}

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) {
    binary += String.fromCharCode(byte)
  }
  return binary
}

function decodeBinaryString(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = cellReferencePattern.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
