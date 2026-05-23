import { MAX_COLS, MAX_ROWS, type WorkbookPreservedPackagePartSnapshot } from '@bilig/protocol'
import { mapStructuralAxisIndex } from '../../engine-structural-utils.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { StructuralAxisTransform } from '@bilig/formula'

const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const binaryChunkSize = 0x8000

export function rewriteControlArtifactsForStructuralTransform(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly transform: StructuralAxisTransform
}): void {
  const sheet = args.workbook.getSheet(args.sheetName)
  const sheetArtifacts = sheet?.controlArtifacts
  if (!sheet || !sheetArtifacts) {
    return
  }

  const nextControlsXml = rewriteControlAnchorXmlForStructuralTransform(sheetArtifacts.controlsXml, args.transform)
  if (nextControlsXml !== sheetArtifacts.controlsXml) {
    sheet.controlArtifacts = { ...sheetArtifacts, controlsXml: nextControlsXml }
  }

  const workbookArtifacts = args.workbook.metadata.controlArtifacts
  if (!workbookArtifacts || workbookArtifacts.parts.length === 0) {
    return
  }

  const vmlPartPaths = controlVmlPartPathsForSheet(args.workbook, args.sheetName)
  if (vmlPartPaths.size === 0) {
    return
  }

  let changed = false
  const parts = workbookArtifacts.parts.map((part) => {
    if (!vmlPartPaths.has(normalizeZipPath(part.path))) {
      return structuredClone(part)
    }
    const nextPart = rewriteVmlPartForStructuralTransform(part, args.transform)
    changed ||= nextPart.dataBase64 !== part.dataBase64 || nextPart.byteLength !== part.byteLength
    return nextPart
  })

  if (changed) {
    args.workbook.metadata.controlArtifacts = {
      ...workbookArtifacts,
      parts,
    }
  }
}

function rewriteControlAnchorXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  if (transform.axis === 'row') {
    return xml.replace(/<((?:[A-Za-z_][\w.-]*:)?row)>(\d+)<\/\1>/gu, (source: string, tagName: string, rowText: string) => {
      const nextRow = rewriteZeroBasedIndexForStructuralTransform(Number(rowText), transform)
      return nextRow === undefined ? source : `<${tagName}>${String(nextRow)}</${tagName}>`
    })
  }
  return xml.replace(/<((?:[A-Za-z_][\w.-]*:)?col)>(\d+)<\/\1>/gu, (source: string, tagName: string, columnText: string) => {
    const nextColumn = rewriteZeroBasedIndexForStructuralTransform(Number(columnText), transform)
    return nextColumn === undefined ? source : `<${tagName}>${String(nextColumn)}</${tagName}>`
  })
}

function rewriteVmlPartForStructuralTransform(
  part: WorkbookPreservedPackagePartSnapshot,
  transform: StructuralAxisTransform,
): WorkbookPreservedPackagePartSnapshot {
  const bytes = decodeBase64(part.dataBase64)
  if (bytes.byteLength !== part.byteLength) {
    return structuredClone(part)
  }
  const xml = new TextDecoder().decode(bytes)
  const nextXml = rewriteVmlAnchorXmlForStructuralTransform(xml, transform)
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

function rewriteVmlAnchorXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  return xml.replace(/<x:Anchor>([\s\S]*?)<\/x:Anchor>/gu, (source: string, body: string) => {
    const values = body
      .split(',')
      .map((value) => Number(value.trim()))
      .filter((value) => Number.isInteger(value))
    if (values.length !== 8) {
      return source
    }

    const firstAxisIndex = transform.axis === 'row' ? 2 : 0
    const secondAxisIndex = transform.axis === 'row' ? 6 : 4
    const first = rewriteZeroBasedIndexForStructuralTransform(values[firstAxisIndex]!, transform)
    const second = rewriteZeroBasedIndexForStructuralTransform(values[secondAxisIndex]!, transform)
    if (first === undefined || second === undefined) {
      return source
    }

    values[firstAxisIndex] = first
    values[secondAxisIndex] = second
    return `<x:Anchor>${values.join(', ')}</x:Anchor>`
  })
}

function rewriteZeroBasedIndexForStructuralTransform(index: number, transform: StructuralAxisTransform): number | undefined {
  const rewritten = mapStructuralAxisIndex(index, transform)
  if (rewritten === undefined) {
    return undefined
  }
  const limit = transform.axis === 'row' ? MAX_ROWS : MAX_COLS
  return rewritten < limit ? rewritten : undefined
}

function controlVmlPartPathsForSheet(workbook: WorkbookStore, sheetName: string): Set<string> {
  const sheet = workbook.getSheet(sheetName)
  const artifacts = sheet?.controlArtifacts
  if (!sheet || !artifacts) {
    return new Set()
  }

  const sheetIndex = [...workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .findIndex((candidate) => candidate.name === sheet.name)
  if (sheetIndex < 0) {
    return new Set()
  }

  const worksheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
  return new Set(
    artifacts.relationships
      .filter((relationship) => relationship.targetMode !== 'External' && relationship.type === vmlDrawingRelationshipType)
      .map((relationship) => resolveTargetPath(worksheetPath, relationship.target)),
  )
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
