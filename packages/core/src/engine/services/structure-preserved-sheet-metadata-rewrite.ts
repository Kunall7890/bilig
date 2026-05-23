import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import type { WorkbookPreservedMetadataRecord, WorkbookPreservedSheetMetadataRecord } from '../../workbook-metadata-types.js'
import { hasPreservedSheetMetadata } from '../../workbook-preserved-metadata.js'

const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu
const pivotTableDefinitionLocationPattern = /<((?:[A-Za-z_][\w.-]*:)?location)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu
const pivotTablePartPathPattern = /^xl\/pivotTables\/pivotTable\d+\.xml$/u
const pivotCacheDefinitionPartPathPattern = /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/u
const worksheetSourcePattern = /<((?:[A-Za-z_][\w.-]*:)?worksheetSource)\b([^>]*\bref=(["'])([^"']+)\3[^>]*)\/?>/gu

export function rewritePreservedSheetMetadataForStructuralTransform(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord | undefined {
  if (!metadata) {
    return undefined
  }
  const next: WorkbookPreservedSheetMetadataRecord = { ...metadata }
  if (metadata.styleArtifacts) {
    const styleArtifacts = rewriteStyleArtifactsForStructuralTransform(metadata.styleArtifacts, transform)
    if (styleArtifacts) {
      next.styleArtifacts = styleArtifacts
    } else {
      delete next.styleArtifacts
    }
  }
  return hasPreservedSheetMetadata(next) ? next : undefined
}

export function rewritePreservedPivotPackageArtifactsForStructuralTransform(
  workbookMetadata: WorkbookPreservedMetadataRecord | undefined,
  sheetMetadata: WorkbookPreservedSheetMetadataRecord | undefined,
  sheetName: string,
  sheetIndex: number,
  transform: StructuralAxisTransform,
): WorkbookPreservedMetadataRecord | undefined {
  const pivotArtifacts = workbookMetadata?.pivotArtifacts
  if (!workbookMetadata || !pivotArtifacts) {
    return workbookMetadata
  }

  const pivotPartPaths = sheetMetadata?.pivotArtifacts ? pivotTablePartPathsForSheet(sheetIndex, sheetMetadata.pivotArtifacts) : new Set()

  return {
    ...workbookMetadata,
    pivotArtifacts: {
      ...pivotArtifacts,
      parts: pivotArtifacts.parts.flatMap((part) => {
        const normalizedPath = normalizePivotPackagePath(part.path)
        if (pivotPartPaths.has(normalizedPath) && pivotTablePartPathPattern.test(normalizedPath)) {
          const xml = rewritePivotTableDefinitionLocationRefsForStructuralTransform(part.xml, transform)
          return xml ? [{ ...part, xml }] : []
        }
        if (pivotCacheDefinitionPartPathPattern.test(normalizedPath)) {
          const xml = rewritePivotCacheWorksheetSourceRefsForStructuralTransform(part.xml, sheetName, transform)
          return xml ? [{ ...part, xml }] : []
        }
        if (!pivotPartPaths.has(normalizedPath)) {
          return [part]
        }
        return [part]
      }),
    },
  }
}

export function preservedSheetMetadataTouchesStructuralDelete(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
  axis: 'row' | 'column',
  start: number,
): boolean {
  if (!metadata) {
    return false
  }
  const styleArtifacts = metadata.styleArtifacts
  if (
    styleArtifacts?.cellStyleIndexes.some((entry) => cellReferenceTouchesAxisDelete(entry.address, axis, start)) ||
    styleArtifacts?.blankCellAddresses?.some((address) => cellReferenceTouchesAxisDelete(address, axis, start))
  ) {
    return true
  }
  return metadata.pivotArtifacts !== undefined
}

function rewriteStyleArtifactsForStructuralTransform(
  styleArtifacts: NonNullable<WorkbookPreservedSheetMetadataRecord['styleArtifacts']>,
  transform: StructuralAxisTransform,
): WorkbookPreservedSheetMetadataRecord['styleArtifacts'] | undefined {
  const cellStyleIndexes = styleArtifacts.cellStyleIndexes.flatMap((entry) => {
    const address = rewriteCellReferenceForStructuralTransform(entry.address, transform)
    return address ? [{ ...entry, address }] : []
  })
  const blankCellAddresses = (styleArtifacts.blankCellAddresses ?? []).flatMap((address) => {
    const nextAddress = rewriteCellReferenceForStructuralTransform(address, transform)
    return nextAddress ? [nextAddress] : []
  })
  if (cellStyleIndexes.length === 0 && blankCellAddresses.length === 0) {
    return undefined
  }
  return {
    cellStyleIndexes,
    ...(blankCellAddresses.length > 0 ? { blankCellAddresses } : {}),
  }
}

function pivotTablePartPathsForSheet(
  sheetIndex: number,
  pivotArtifacts: NonNullable<WorkbookPreservedSheetMetadataRecord['pivotArtifacts']>,
): Set<string> {
  const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
  return new Set(
    pivotArtifacts.relationships
      .filter((relationship) => relationship.type.endsWith('/pivotTable'))
      .map((relationship) => normalizePivotPackagePath(resolvePivotRelationshipTarget(sheetPath, relationship.target))),
  )
}

function resolvePivotRelationshipTarget(basePartPath: string, target: string): string {
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

function normalizePivotPackagePath(path: string): string {
  return path.replace(/^\/+/u, '').replace(/\\/gu, '/')
}

function rewritePivotTableDefinitionLocationRefsForStructuralTransform(
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

function rewritePivotCacheWorksheetSourceRefsForStructuralTransform(
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

function readXmlAttribute(attributes: string, name: string): string | undefined {
  return new RegExp(`\\b${name}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2]
}

function unescapeXmlAttribute(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}

function rewriteCellReferenceForStructuralTransform(address: string, transform: StructuralAxisTransform): string | undefined {
  if (!isCellReference(address)) {
    return address
  }
  const rewritten = rewriteAddressForStructuralTransform(address, transform)
  return rewritten ? clipAddressToSheetGrid(rewritten) : undefined
}

function cellReferenceTouchesAxisDelete(address: string, axis: 'row' | 'column', start: number): boolean {
  if (!isCellReference(address)) {
    return false
  }
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return false
  }
  return (axis === 'row' ? parsed[0] : parsed[1]) >= start
}

function isCellReference(address: string): boolean {
  return cellReferencePattern.test(address)
}

function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = cellReferencePattern.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
