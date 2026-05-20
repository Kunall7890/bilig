import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookDataModelArtifactsSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import {
  getZipText,
  normalizeZipPath,
  readXlsxZipEntries,
  readXlsxZipEntryUncompressedSize,
  type XlsxZipEntries,
  type XlsxZipSource,
} from './xlsx-zip.js'
import {
  buildRelationshipsXml,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'
import { decodedPartBytes, encodedPartSnapshot, lazyEncodedPartSnapshot } from './xlsx-preserved-package-parts.js'

const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const contentTypesPath = '[Content_Types].xml'
const powerPivotDataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData'
const customXmlRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
const customDataPropertiesRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/customDataProps'
const dataModelPackagePartPattern = /^xl\/model\//u
const customDataPackagePartPattern = /^xl\/customData\//u
const customXmlPackagePartPattern = /^customXml\//u

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
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
  const pattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${escapeRegExp(extension)}\\1`, 'u')
  if (pattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/></Types>`,
  )
}

function upsertContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (!contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  const escapedPartName = escapeXml(partName)
  const escapedContentType = escapeXml(contentType)
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  let replaced = false
  const nextXml = contentTypesXml.replace(overridePattern, (match: string, attributes: string) => {
    if (readAttribute(attributes, 'PartName') !== partName) {
      return match
    }
    replaced = true
    return `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/>`
  })
  return replaced
    ? nextXml
    : nextXml.replace('</Types>', `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/></Types>`)
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

function isDataModelRelationship(relationship: ParsedRelationship): boolean {
  return (
    relationship.type === powerPivotDataRelationshipType ||
    relationship.type === customXmlRelationshipType ||
    relationship.type === customDataPropertiesRelationshipType
  )
}

function isDataModelPackagePartPath(path: string): boolean {
  return dataModelPackagePartPattern.test(path) || customDataPackagePartPattern.test(path) || customXmlPackagePartPattern.test(path)
}

function preservedPartsByPath(parts: readonly WorkbookPreservedPackagePartSnapshot[]): Map<string, Uint8Array> {
  const output = new Map<string, Uint8Array>()
  for (const part of parts) {
    const bytes = decodedPartBytes(part)
    if (bytes) {
      output.set(normalizeZipPath(part.path), bytes)
    }
  }
  return output
}

function relationshipTargetExists(
  relationship: WorkbookPackageRelationshipSnapshot,
  partsByPath: ReadonlyMap<string, Uint8Array>,
  zip: XlsxZipEntries,
): boolean {
  const targetPath = normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))
  return partsByPath.has(targetPath) || Boolean(zip[targetPath])
}

function addWorkbookRelationships(
  relationships: ParsedRelationship[],
  preservedRelationships: readonly WorkbookPackageRelationshipSnapshot[],
  partsByPath: ReadonlyMap<string, Uint8Array>,
  zip: XlsxZipEntries,
): boolean {
  let changed = false
  for (const preservedRelationship of preservedRelationships) {
    if (!relationshipTargetExists(preservedRelationship, partsByPath, zip)) {
      continue
    }
    const existing = relationships.find(
      (relationship) => relationship.type === preservedRelationship.type && relationship.target === preservedRelationship.target,
    )
    if (existing) {
      continue
    }
    const idInUse = relationships.some((relationship) => relationship.id === preservedRelationship.id)
    relationships.push({
      ...parsedRelationship(preservedRelationship),
      id: preservedRelationship.id.length > 0 && !idInUse ? preservedRelationship.id : nextRelationshipId(relationships),
    })
    changed = true
  }
  return changed
}

export function readImportedWorkbookDataModelArtifacts(source: XlsxZipSource): WorkbookDataModelArtifactsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const partPaths = Object.keys(zip).filter(isDataModelPackagePartPath).toSorted()
  if (partPaths.length === 0) {
    return undefined
  }

  const relationships = parseRelationships(getZipText(zip, workbookRelationshipsPath)).filter(isDataModelRelationship)
  const contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const parts = partPaths.flatMap((path) => {
    const byteLength = readXlsxZipEntryUncompressedSize(zip, path)
    if (byteLength !== undefined) {
      return [
        lazyEncodedPartSnapshot(path, byteLength, () => {
          const bytes = zip[path]
          if (bytes) {
            Reflect.deleteProperty(zip, path)
          }
          return bytes
        }),
      ]
    }
    const bytes = zip[path]
    if (!bytes) {
      return []
    }
    const snapshot = encodedPartSnapshot(path, bytes)
    Reflect.deleteProperty(zip, path)
    return [snapshot]
  })
  return {
    parts,
    workbookRelationships: relationships.map(relationshipSnapshot),
    ...(contentTypesXml
      ? {
          ...(readContentTypeDefaults(contentTypesXml, partPaths).length > 0
            ? { contentTypeDefaults: readContentTypeDefaults(contentTypesXml, partPaths) }
            : {}),
          ...(readContentTypeOverrides(contentTypesXml, partPaths).length > 0
            ? { contentTypeOverrides: readContentTypeOverrides(contentTypesXml, partPaths) }
            : {}),
        }
      : {}),
  }
}

export function addExportDataModelArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const artifacts = snapshot.workbook.metadata?.dataModelArtifacts
  if (!artifacts || artifacts.parts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(artifacts.parts)
  if (partsByPath.size === 0) {
    return bytes
  }

  const copiedPartPaths = new Set<string>()
  for (const [path, partBytes] of partsByPath) {
    zip[path] = partBytes
    copiedPartPaths.add(path)
  }

  const relationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const relationshipsChanged = addWorkbookRelationships(relationships, artifacts.workbookRelationships, partsByPath, zip)
  if (relationshipsChanged) {
    setZipText(zip, workbookRelationshipsPath, buildRelationshipsXml(relationships))
  }

  let contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const copiedExtensions = new Set(
    [...copiedPartPaths].map(extensionFromPath).filter((extension): extension is string => Boolean(extension)),
  )
  for (const defaultEntry of artifacts.contentTypeDefaults ?? []) {
    if (copiedExtensions.has(defaultEntry.extension)) {
      contentTypesXml = addContentTypeDefault(contentTypesXml, defaultEntry.extension, defaultEntry.contentType)
    }
  }
  for (const overrideEntry of artifacts.contentTypeOverrides ?? []) {
    if (copiedPartPaths.has(normalizeZipPath(overrideEntry.partName))) {
      contentTypesXml = upsertContentTypeOverride(contentTypesXml, overrideEntry.partName, overrideEntry.contentType)
    }
  }
  if (contentTypesXml.length > 0) {
    setZipText(zip, contentTypesPath, contentTypesXml)
  }

  return zipSync(zip)
}
