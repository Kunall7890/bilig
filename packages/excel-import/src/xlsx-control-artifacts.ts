import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookControlArtifactsSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSheetControlArtifactsSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'

const binaryChunkSize = 0x8000
const controlRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const namespaceDeclarationPattern = /\s(xmlns(?::[A-Za-z_][\w.-]*)?)=("|')([\s\S]*?)\2/gu
const controlsTailElements = ['webPublishItems', 'tableParts', 'extLst'] as const
const legacyDrawingTailElements = ['drawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', ...controlsTailElements] as const

interface ElementRange {
  readonly start: number
  readonly end: number
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
  const btoa = globalThis.btoa
  if (typeof btoa === 'function') {
    return btoa(encodeBinaryString(bytes))
  }
  return Buffer.from(bytes).toString('base64')
}

function decodeBase64(dataBase64: string): Uint8Array {
  const atob = globalThis.atob
  if (typeof atob === 'function') {
    return decodeBinaryString(atob(dataBase64))
  }
  return new Uint8Array(Buffer.from(dataBase64, 'base64'))
}

function encodedPartSnapshot(path: string, bytes: Uint8Array): WorkbookPreservedPackagePartSnapshot {
  return {
    path,
    storage: 'base64',
    dataBase64: encodeBase64(bytes),
    byteLength: bytes.byteLength,
  }
}

function decodedPartBytes(part: WorkbookPreservedPackagePartSnapshot): Uint8Array | undefined {
  if (part.storage !== 'base64') {
    return undefined
  }
  const bytes = decodeBase64(part.dataBase64)
  return bytes.byteLength === part.byteLength ? bytes : undefined
}

function preservedPartsByPath(parts: readonly WorkbookPreservedPackagePartSnapshot[]): Map<string, Uint8Array> {
  const output = new Map<string, Uint8Array>()
  parts.forEach((part) => {
    const bytes = decodedPartBytes(part)
    if (bytes) {
      output.set(normalizeZipPath(part.path), bytes)
    }
  })
  return output
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function elementRangesByLocalName(xml: string, localName: string): ElementRange[] {
  const tagPattern = new RegExp(`<(/?)(?:[A-Za-z_][\\w.-]*:)?${escapeRegExp(localName)}\\b[^>]*(/?)>`, 'gu')
  const ranges: ElementRange[] = []
  const stack: number[] = []
  for (const match of xml.matchAll(tagPattern)) {
    const tagXml = match[0]
    const start = match.index
    const isClosing = match[1] === '/'
    const isSelfClosing = match[2] === '/' || tagXml.endsWith('/>')
    if (isClosing) {
      const openStart = stack.pop()
      if (openStart !== undefined) {
        ranges.push({ start: openStart, end: start + tagXml.length })
      }
      continue
    }
    if (isSelfClosing) {
      ranges.push({ start, end: start + tagXml.length })
      continue
    }
    stack.push(start)
  }
  return ranges.toSorted((left, right) => left.start - right.start || left.end - right.end)
}

function controlsFragmentRanges(sheetXml: string): ElementRange[] {
  const alternateContentRanges = elementRangesByLocalName(sheetXml, 'AlternateContent')
  const ranges = elementRangesByLocalName(sheetXml, 'controls').map((controlsRange) => {
    const enclosingAlternate = alternateContentRanges
      .filter((range) => range.start <= controlsRange.start && range.end >= controlsRange.end)
      .toSorted((left, right) => left.end - left.start - (right.end - right.start))[0]
    return enclosingAlternate ?? controlsRange
  })
  const deduped: ElementRange[] = []
  for (const range of ranges.toSorted((left, right) => left.start - right.start || left.end - right.end)) {
    if (deduped.some((entry) => entry.start === range.start && entry.end === range.end)) {
      continue
    }
    deduped.push(range)
  }
  return deduped
}

function readControlsXml(sheetXml: string | null): string | undefined {
  if (!sheetXml) {
    return undefined
  }
  const ranges = controlsFragmentRanges(sheetXml)
  return ranges.length > 0 ? ranges.map((range) => sheetXml.slice(range.start, range.end)).join('') : undefined
}

function removeControlsXml(sheetXml: string): string {
  let output = sheetXml
  for (const range of controlsFragmentRanges(sheetXml).toReversed()) {
    output = `${output.slice(0, range.start)}${output.slice(range.end)}`
  }
  return output
}

function readWorksheetRootOpenTag(sheetXml: string | null): string | undefined {
  return sheetXml ? /<worksheet\b[^>]*>/u.exec(sheetXml)?.[0] : undefined
}

function readLegacyDrawingRelationshipId(sheetXml: string | null): string | null {
  const legacyDrawingTag = /<legacyDrawing\b[^>]*(?:\/?>|>[\s\S]*?<\/legacyDrawing>)/u.exec(sheetXml ?? '')?.[0]
  return legacyDrawingTag ? (readAttribute(legacyDrawingTag, 'r:id') ?? readAttribute(legacyDrawingTag, 'id')) : null
}

function relationshipSnapshot(relationship: ParsedRelationship): WorkbookPackageRelationshipSnapshot {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function parsedRelationship(relationship: WorkbookPackageRelationshipSnapshot): ParsedRelationship {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function addPreservedRelationships(
  relationships: ParsedRelationship[],
  preservedRelationships: readonly WorkbookPackageRelationshipSnapshot[],
): Map<string, string> {
  const relationshipIds = new Map<string, string>()
  for (const preservedRelationship of preservedRelationships) {
    const existing = relationships.find(
      (relationship) => relationship.type === preservedRelationship.type && relationship.target === preservedRelationship.target,
    )
    if (existing) {
      relationshipIds.set(preservedRelationship.id, existing.id)
      continue
    }
    const idInUse = relationships.some((relationship) => relationship.id === preservedRelationship.id)
    const nextId = preservedRelationship.id.length > 0 && !idInUse ? preservedRelationship.id : nextRelationshipId(relationships)
    relationships.push({ ...parsedRelationship(preservedRelationship), id: nextId })
    relationshipIds.set(preservedRelationship.id, nextId)
  }
  return relationshipIds
}

function replaceControlRelationshipIds(xml: string, relationshipIds: ReadonlyMap<string, string>): string {
  if (relationshipIds.size === 0) {
    return xml
  }
  return xml.replace(/\br:id=(["'])([\s\S]*?)\1/gu, (match, quote: string, id: string) => {
    const nextId = relationshipIds.get(id)
    return nextId ? match.replace(`${quote}${id}${quote}`, `${quote}${nextId}${quote}`) : match
  })
}

function addXmlAttribute(openTag: string, name: string, value: string): string {
  return openTag.replace(/>$/u, ` ${name}="${escapeXml(value)}">`)
}

function mergeIgnorableAttribute(targetRootOpenTag: string, sourceRootOpenTag: string): string {
  const sourceIgnorable = readAttribute(sourceRootOpenTag, 'mc:Ignorable')
  if (!sourceIgnorable) {
    return targetRootOpenTag
  }
  const targetIgnorable = readAttribute(targetRootOpenTag, 'mc:Ignorable')
  if (!targetIgnorable) {
    return addXmlAttribute(targetRootOpenTag, 'mc:Ignorable', sourceIgnorable)
  }
  const nextTokens = [...new Set([...targetIgnorable.split(/\s+/u), ...sourceIgnorable.split(/\s+/u)].filter(Boolean))]
  return targetRootOpenTag.replace(/\smc:Ignorable=(["'])([\s\S]*?)\1/u, ` mc:Ignorable="${escapeXml(nextTokens.join(' '))}"`)
}

function mergeWorksheetRootOpenTag(targetRootOpenTag: string, sourceRootOpenTag: string): string {
  let output = targetRootOpenTag
  namespaceDeclarationPattern.lastIndex = 0
  for (const match of sourceRootOpenTag.matchAll(namespaceDeclarationPattern)) {
    const name = match[1]
    const quote = match[2]
    const value = match[3]
    if (!name || !quote || value === undefined || new RegExp(`\\s${escapeRegExp(name)}=("|')`, 'u').test(output)) {
      continue
    }
    output = output.replace(/>$/u, ` ${name}=${quote}${escapeXml(value)}${quote}>`)
  }
  return mergeIgnorableAttribute(output, sourceRootOpenTag)
}

function mergeWorksheetControlNamespaces(sheetXml: string, sourceRootOpenTag: string): string {
  return sheetXml.replace(/<worksheet\b[^>]*>/u, (targetRootOpenTag) => mergeWorksheetRootOpenTag(targetRootOpenTag, sourceRootOpenTag))
}

function insertBeforeTailElement(sheetXml: string, xml: string, tailElements: readonly string[]): string {
  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of tailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${xml}${sheetXml.slice(insertIndex)}`
}

function upsertLegacyDrawing(sheetXml: string, relationshipId: string): string {
  const withNamespace = ensureRelationshipNamespace(sheetXml)
  const legacyDrawingXml = `<legacyDrawing r:id="${escapeXml(relationshipId)}"/>`
  if (/<legacyDrawing\b[^>]*(?:\/>|>[\s\S]*?<\/legacyDrawing>)/u.test(withNamespace)) {
    return withNamespace.replace(/<legacyDrawing\b[^>]*(?:\/>|>[\s\S]*?<\/legacyDrawing>)/u, legacyDrawingXml)
  }
  return insertBeforeTailElement(withNamespace, legacyDrawingXml, legacyDrawingTailElements)
}

function upsertControlsXml(sheetXml: string, controlsXml: string, worksheetRootOpenTag: string): string {
  const withoutControls = removeControlsXml(mergeWorksheetControlNamespaces(sheetXml, worksheetRootOpenTag))
  return insertBeforeTailElement(withoutControls, controlsXml, controlsTailElements)
}

function extensionFromPath(path: string): string | null {
  const fileName = normalizeZipPath(path).slice(normalizeZipPath(path).lastIndexOf('/') + 1)
  const extensionIndex = fileName.lastIndexOf('.')
  return extensionIndex >= 0 && extensionIndex < fileName.length - 1 ? fileName.slice(extensionIndex + 1).toLowerCase() : null
}

function readContentTypeDefaults(contentTypesXml: string, partPaths: readonly string[]): WorkbookContentTypeDefaultSnapshot[] {
  const neededExtensions = new Set(partPaths.map(extensionFromPath).filter((extension): extension is string => Boolean(extension)))
  const defaultsByExtension = new Map<string, WorkbookContentTypeDefaultSnapshot>()
  for (const match of contentTypesXml.matchAll(/<Default\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    const extension = readAttribute(attributes, 'Extension')?.toLowerCase()
    const contentType = readAttribute(attributes, 'ContentType')
    if (!extension || !contentType || !neededExtensions.has(extension)) {
      continue
    }
    defaultsByExtension.set(extension, { extension, contentType })
  }
  return [...defaultsByExtension.values()].toSorted((left, right) => left.extension.localeCompare(right.extension))
}

function readContentTypeOverrides(contentTypesXml: string, partPaths: readonly string[]): WorkbookContentTypeOverrideSnapshot[] {
  const neededPartNames = new Set(partPaths.map((path) => `/${normalizeZipPath(path)}`))
  const overridesByPartName = new Map<string, WorkbookContentTypeOverrideSnapshot>()
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    const partName = readAttribute(attributes, 'PartName')
    const contentType = readAttribute(attributes, 'ContentType')
    if (!partName || !contentType || !neededPartNames.has(partName)) {
      continue
    }
    overridesByPartName.set(partName, { partName, contentType })
  }
  return [...overridesByPartName.values()].toSorted((left, right) => left.partName.localeCompare(right.partName))
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  const existingDefaultPattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${escapeRegExp(extension)}\\1`, 'u')
  if (existingDefaultPattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/></Types>`,
  )
}

function copyRelationshipPart(input: {
  readonly zip: XlsxZipEntries
  readonly sheetPath: string
  readonly relationship: WorkbookPackageRelationshipSnapshot
  readonly partsByPath: ReadonlyMap<string, Uint8Array>
  readonly copiedPartPaths: Set<string>
}): boolean {
  const partPath = normalizeZipPath(resolveTargetPath(input.sheetPath, input.relationship.target))
  const bytes = input.partsByPath.get(partPath)
  if (!bytes) {
    return false
  }
  input.zip[partPath] = bytes
  input.copiedPartPaths.add(partPath)
  return true
}

export function readImportedWorkbookControlArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): {
  readonly artifacts: WorkbookControlArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetControlArtifactsSnapshot>
} {
  const zip = readXlsxZipEntries(source)
  const sheetArtifactsByName = new Map<string, WorkbookSheetControlArtifactsSnapshot>()
  const partPaths = new Set<string>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    const controlsXml = readControlsXml(sheetXml)
    const worksheetRootOpenTag = readWorksheetRootOpenTag(sheetXml)
    if (!controlsXml || !worksheetRootOpenTag) {
      return
    }

    const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
    const relationshipIds = new Set(
      [...controlsXml.matchAll(/\br:id=(["'])([\s\S]*?)\1/gu)].map((match) => match[2]).filter((id): id is string => Boolean(id)),
    )
    const legacyDrawingRelationshipId = readLegacyDrawingRelationshipId(sheetXml)
    if (legacyDrawingRelationshipId) {
      relationshipIds.add(legacyDrawingRelationshipId)
    }
    const relationships = sheetRelationships.filter(
      (relationship) =>
        relationshipIds.has(relationship.id) &&
        (relationship.type === controlRelationshipType || relationship.type === vmlDrawingRelationshipType),
    )
    if (relationships.length === 0) {
      return
    }

    sheetArtifactsByName.set(sheetName, {
      controlsXml,
      worksheetRootOpenTag,
      relationships: relationships.map(relationshipSnapshot),
    })
    relationships.forEach((relationship) => {
      const path = normalizeZipPath(resolveTargetPath(sheetPath, relationship.target))
      if (zip[path]) {
        partPaths.add(path)
      }
    })
  })

  if (partPaths.size === 0) {
    return {
      artifacts: undefined,
      sheetArtifactsByName,
    }
  }

  const contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  const parts = [...partPaths].toSorted().flatMap((path) => {
    const bytes = zip[path]
    return bytes ? [encodedPartSnapshot(path, bytes)] : []
  })

  return {
    artifacts: {
      parts,
      ...(contentTypesXml
        ? {
            ...(readContentTypeDefaults(contentTypesXml, [...partPaths]).length > 0
              ? { contentTypeDefaults: readContentTypeDefaults(contentTypesXml, [...partPaths]) }
              : {}),
            ...(readContentTypeOverrides(contentTypesXml, [...partPaths]).length > 0
              ? { contentTypeOverrides: readContentTypeOverrides(contentTypesXml, [...partPaths]) }
              : {}),
          }
        : {}),
    },
    sheetArtifactsByName,
  }
}

export function addExportControlArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookArtifacts = snapshot.workbook.metadata?.controlArtifacts
  const sheetsWithControlArtifacts = snapshot.sheets.filter((sheet) => sheet.metadata?.controlArtifacts)
  if (!workbookArtifacts || sheetsWithControlArtifacts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(workbookArtifacts.parts)
  const copiedPartPaths = new Set<string>()
  let changed = false

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetControlArtifacts = sheet.metadata?.controlArtifacts
      if (!sheetControlArtifacts) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }

      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))
      const relationshipIds = addPreservedRelationships(sheetRelationships, sheetControlArtifacts.relationships)
      let nextSheetXml = sheetXml
      const legacyDrawingRelationship = sheetControlArtifacts.relationships.find(
        (relationship) => relationship.type === vmlDrawingRelationshipType,
      )
      const nextLegacyDrawingRelationshipId = legacyDrawingRelationship ? relationshipIds.get(legacyDrawingRelationship.id) : undefined
      if (nextLegacyDrawingRelationshipId) {
        nextSheetXml = upsertLegacyDrawing(nextSheetXml, nextLegacyDrawingRelationshipId)
      }
      nextSheetXml = upsertControlsXml(
        nextSheetXml,
        replaceControlRelationshipIds(sheetControlArtifacts.controlsXml, relationshipIds),
        sheetControlArtifacts.worksheetRootOpenTag,
      )

      sheetControlArtifacts.relationships.forEach((relationship) => {
        changed = copyRelationshipPart({ zip, sheetPath, relationship, partsByPath, copiedPartPaths }) || changed
      })
      setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      setZipText(zip, sheetPath, nextSheetXml)
      changed = true
    })

  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  const copiedExtensions = new Set(
    [...copiedPartPaths].map(extensionFromPath).filter((extension): extension is string => Boolean(extension)),
  )
  for (const defaultEntry of workbookArtifacts.contentTypeDefaults ?? []) {
    if (!copiedExtensions.has(defaultEntry.extension)) {
      continue
    }
    contentTypesXml = addContentTypeDefault(contentTypesXml, defaultEntry.extension, defaultEntry.contentType)
  }
  for (const overrideEntry of workbookArtifacts.contentTypeOverrides ?? []) {
    if (!copiedPartPaths.has(normalizeZipPath(overrideEntry.partName))) {
      continue
    }
    contentTypesXml = addContentTypeOverride(contentTypesXml, overrideEntry.partName, overrideEntry.contentType)
  }
  if (copiedExtensions.has('vml')) {
    contentTypesXml = addContentTypeDefault(contentTypesXml, 'vml', vmlDrawingContentType)
  }
  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }

  return changed ? zipSync(zip) : bytes
}
