import { MAX_COLS, MAX_ROWS, type WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import { mapStructuralAxisIndex } from '../../engine-structural-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'

const binaryChunkSize = 0x8000

export function rewriteDrawingArtifactsForStructuralTransform(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly transform: StructuralAxisTransform
}): void {
  const sheetArtifacts = args.workbook.getSheetDrawingArtifacts(args.sheetName)
  const workbookArtifacts = args.workbook.getDrawingArtifacts()
  if (!sheetArtifacts || !workbookArtifacts || workbookArtifacts.parts.length === 0) {
    return
  }

  const drawingPartPaths = drawingPartPathsForSheet(args.workbook, args.sheetName)
  if (drawingPartPaths.size === 0) {
    return
  }

  let changed = false
  const parts = workbookArtifacts.parts.map((part) => {
    if (!drawingPartPaths.has(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteDrawingPartForStructuralTransform(part, args.transform)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  if (changed) {
    args.workbook.setDrawingArtifacts({ ...workbookArtifacts, parts })
  }
}

export function drawingArtifactsTouchStructuralDelete(workbook: WorkbookStore, sheetName: string): boolean {
  return workbook.getSheetDrawingArtifacts(sheetName) !== undefined && workbook.getDrawingArtifacts() !== undefined
}

function drawingPartPathsForSheet(workbook: WorkbookStore, sheetName: string): Set<string> {
  const sheetArtifacts = workbook.getSheetDrawingArtifacts(sheetName)
  if (!sheetArtifacts) {
    return new Set()
  }

  const sheetIndex = [...workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .findIndex((sheet) => sheet.name === sheetName)
  if (sheetIndex < 0) {
    return new Set()
  }

  const worksheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
  return new Set([resolveTargetPath(worksheetPath, sheetArtifacts.relationshipTarget)])
}

function rewriteDrawingPartForStructuralTransform(
  part: WorkbookPreservedPackagePartSnapshot,
  transform: StructuralAxisTransform,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = rewriteDrawingAnchorXmlForStructuralTransform(xml, transform)
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

function rewriteDrawingAnchorXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  return xml.replace(
    /<((?:[A-Za-z_][\w.-]*:)?(?:from|to))\b([^>]*)>([\s\S]*?)<\/\1>/gu,
    (source: string, tagName: string, attributes: string, body: string) => {
      const nextBody = rewriteDrawingMarkerBodyForStructuralTransform(body, transform)
      return nextBody === body ? source : `<${tagName}${attributes}>${nextBody}</${tagName}>`
    },
  )
}

function rewriteDrawingMarkerBodyForStructuralTransform(body: string, transform: StructuralAxisTransform): string {
  const axisTagName = transform.axis === 'row' ? 'row' : 'col'
  return body.replace(
    new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${axisTagName})>(\\d+)</\\1>`, 'gu'),
    (source: string, tagName: string, indexText: string) => {
      const index = Number(indexText)
      if (!Number.isSafeInteger(index)) {
        return source
      }
      const nextIndex = mapDrawingMarkerIndexForStructuralTransform(index, transform)
      return nextIndex === undefined ? source : `<${tagName}>${String(nextIndex)}</${tagName}>`
    },
  )
}

function mapDrawingMarkerIndexForStructuralTransform(index: number, transform: StructuralAxisTransform): number | undefined {
  const rewritten = mapStructuralAxisIndex(index, transform)
  if (rewritten === undefined) {
    return undefined
  }
  const limit = transform.axis === 'row' ? MAX_ROWS : MAX_COLS
  return rewritten < limit ? rewritten : undefined
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
