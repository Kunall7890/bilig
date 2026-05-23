import { MAX_COLS, MAX_ROWS, type WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import { columnToIndex, formatAddress, rewriteAddressForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'
import type { WorkbookSheetThreadedCommentArtifactsRecord, WorkbookThreadedCommentArtifactsRecord } from '../../workbook-metadata-types.js'
import type { WorkbookStore } from '../../workbook-store.js'

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i
const binaryChunkSize = 0x8000
const workbookPath = 'xl/workbook.xml'
const worksheetPathForRelationshipResolution = 'xl/worksheets/sheet1.xml'
const threadedCommentRelationshipTypeFragment = '/relationships/threadedComment'
const personRelationshipTypeFragment = '/relationships/person'
const threadedCommentPartPathPattern = /^xl\/threadedComments\/threadedComment[^/]*\.xml$/u
const personPartPathPattern = /^xl\/persons\/person[^/]*\.xml$/u

export function rewriteThreadedCommentArtifactsForStructuralTransform(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly transform: StructuralAxisTransform
}): void {
  const sheetArtifacts = args.workbook.getSheetThreadedCommentArtifacts(args.sheetName)
  const workbookArtifacts = args.workbook.getThreadedCommentArtifacts()
  if (!sheetArtifacts || !workbookArtifacts || workbookArtifacts.parts.length === 0) {
    return
  }

  const sheetIndex = [...args.workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .findIndex((sheet) => sheet.name === args.sheetName)
  if (sheetIndex < 0) {
    return
  }

  const worksheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
  const threadedCommentPartPaths = new Set(
    sheetArtifacts.relationships
      .filter((relationship) => relationship.targetMode !== 'External')
      .map((relationship) => resolveTargetPath(worksheetPath, relationship.target)),
  )
  if (threadedCommentPartPaths.size === 0) {
    return
  }

  let changed = false
  const parts = workbookArtifacts.parts.map((part) => {
    if (!threadedCommentPartPaths.has(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteThreadedCommentPartForStructuralTransform(part, args.transform)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  if (changed) {
    args.workbook.setThreadedCommentArtifacts({
      ...workbookArtifacts,
      parts,
    })
  }
}

export function rewriteThreadedCommentArtifactsForSheetDeletion(args: {
  readonly workbookArtifacts: WorkbookThreadedCommentArtifactsRecord | undefined
  readonly sheetArtifactsByName: ReadonlyMap<string, WorkbookSheetThreadedCommentArtifactsRecord>
  readonly deletedSheetName: string
}): WorkbookThreadedCommentArtifactsRecord | undefined {
  const workbookArtifacts = args.workbookArtifacts
  const deletedSheetArtifacts = args.sheetArtifactsByName.get(args.deletedSheetName)
  if (!workbookArtifacts || !deletedSheetArtifacts) {
    return workbookArtifacts
  }

  const survivingThreadedCommentPartPaths = new Set<string>()
  for (const [sheetName, sheetArtifacts] of args.sheetArtifactsByName.entries()) {
    if (sheetName === args.deletedSheetName) {
      continue
    }
    for (const path of threadedCommentPartPathsReferencedBySheetArtifacts(sheetArtifacts)) {
      survivingThreadedCommentPartPaths.add(path)
    }
  }

  const retainedThreadedCommentParts: WorkbookPreservedPackagePartSnapshot[] = []
  const removedPartPaths = new Set<string>()
  for (const part of workbookArtifacts.parts) {
    const path = normalizeZipPath(part.path)
    if (!threadedCommentPartPathPattern.test(path)) {
      continue
    }
    if (!survivingThreadedCommentPartPaths.has(path)) {
      removedPartPaths.add(path)
      continue
    }
    retainedThreadedCommentParts.push(structuredClone(part))
  }

  const retainedPersonParts = workbookArtifacts.parts.flatMap((part) => {
    const path = normalizeZipPath(part.path)
    if (!personPartPathPattern.test(path)) {
      return []
    }
    return [retainedThreadedCommentParts.length === 0 ? clearPersonPartEntries(part) : structuredClone(part)]
  })

  const retainedOtherParts = workbookArtifacts.parts.flatMap((part) => {
    const path = normalizeZipPath(part.path)
    if (threadedCommentPartPathPattern.test(path) || personPartPathPattern.test(path)) {
      return []
    }
    return [structuredClone(part)]
  })

  const parts = [...retainedThreadedCommentParts, ...retainedPersonParts, ...retainedOtherParts].toSorted((left, right) =>
    normalizeZipPath(left.path).localeCompare(normalizeZipPath(right.path)),
  )
  const retainedPartPaths = new Set(parts.map((part) => normalizeZipPath(part.path)))
  const workbookRelationships = (workbookArtifacts.workbookRelationships ?? []).filter((relationship) => {
    if (relationship.targetMode === 'External') {
      return true
    }
    const targetPath = resolveTargetPath(workbookPath, relationship.target)
    if (relationship.type.includes(personRelationshipTypeFragment) || personPartPathPattern.test(targetPath)) {
      return retainedPartPaths.has(targetPath)
    }
    return true
  })
  const contentTypeOverrides = (workbookArtifacts.contentTypeOverrides ?? []).filter((entry) => {
    const path = normalizeZipPath(entry.partName)
    if (threadedCommentPartPathPattern.test(path) || personPartPathPattern.test(path)) {
      return retainedPartPaths.has(path)
    }
    return !removedPartPaths.has(path)
  })

  const next: WorkbookThreadedCommentArtifactsRecord = {
    parts,
    ...(workbookRelationships.length > 0 ? { workbookRelationships } : {}),
    ...(parts.length > 0 && workbookArtifacts.contentTypeDefaults
      ? { contentTypeDefaults: structuredClone(workbookArtifacts.contentTypeDefaults) }
      : {}),
    ...(contentTypeOverrides.length > 0 ? { contentTypeOverrides } : {}),
  }
  return hasThreadedCommentArtifacts(next) ? next : undefined
}

function rewriteThreadedCommentPartForStructuralTransform(
  part: WorkbookPreservedPackagePartSnapshot,
  transform: StructuralAxisTransform,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = rewriteThreadedCommentXmlForStructuralTransform(xml, transform)
  if (nextXml === xml) {
    return structuredClone(part)
  }
  const nextBytes = new TextEncoder().encode(nextXml)
  return {
    ...part,
    dataBase64: encodeBase64(nextBytes),
    byteLength: nextBytes.byteLength,
  }
}

function threadedCommentPartPathsReferencedBySheetArtifacts(artifacts: WorkbookSheetThreadedCommentArtifactsRecord): Set<string> {
  return new Set(
    artifacts.relationships
      .filter((relationship) => relationship.targetMode !== 'External')
      .filter((relationship) => relationship.type.includes(threadedCommentRelationshipTypeFragment))
      .map((relationship) => resolveTargetPath(worksheetPathForRelationshipResolution, relationship.target)),
  )
}

function clearPersonPartEntries(part: WorkbookPreservedPackagePartSnapshot): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = xml.replace(/<personList\b([^>]*)>[\s\S]*?<\/personList>/u, '<personList$1/>')
  if (nextXml === xml) {
    return structuredClone(part)
  }
  const nextBytes = new TextEncoder().encode(nextXml)
  return {
    ...part,
    dataBase64: encodeBase64(nextBytes),
    byteLength: nextBytes.byteLength,
  }
}

function hasThreadedCommentArtifacts(artifacts: WorkbookThreadedCommentArtifactsRecord): boolean {
  return (
    artifacts.parts.length > 0 || (artifacts.workbookRelationships?.length ?? 0) > 0 || (artifacts.contentTypeOverrides?.length ?? 0) > 0
  )
}

function rewriteThreadedCommentXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  return xml.replace(/<threadedComment\b([^>]*?)(\/>|>[\s\S]*?<\/threadedComment>)/gu, (source: string, attributes: string) => {
    const ref = readXmlAttribute(attributes, 'ref')
    if (!ref) {
      return source
    }
    const nextRef = rewriteThreadedCommentRefForStructuralTransform(ref, transform)
    if (!nextRef) {
      return ''
    }
    return source.replace(/\bref=(["'])([\s\S]*?)\1/u, (_attribute: string, quote: string) => `ref=${quote}${nextRef}${quote}`)
  })
}

function rewriteThreadedCommentRefForStructuralTransform(ref: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(ref, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseCellAddress(rewritten)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = METADATA_CELL_REF_RE.exec(address)
  return match ? [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())] : undefined
}

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function normalizeZipPath(path: string): string {
  return path.replace(/\\/gu, '/').replace(/^\/+/u, '')
}

function resolveTargetPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const baseSegments = normalizeZipPath(basePartPath).split('/')
  baseSegments.pop()
  for (const segment of target.split('/')) {
    if (!segment || segment === '.') {
      continue
    }
    if (segment === '..') {
      baseSegments.pop()
    } else {
      baseSegments.push(segment)
    }
  }
  return normalizeZipPath(baseSegments.join('/'))
}

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
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

function encodeBase64(bytes: Uint8Array): string {
  return globalThis.btoa(encodeBinaryString(bytes))
}

function decodeBase64(dataBase64: string): Uint8Array {
  return decodeBinaryString(globalThis.atob(dataBase64))
}
