import { MAX_COLS, MAX_ROWS, type WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import { columnToIndex, formatAddress, rewriteAddressForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'
import type { WorkbookStore } from '../../workbook-store.js'

const METADATA_CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i
const binaryChunkSize = 0x8000

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
