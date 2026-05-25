import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookExternalLinkArtifactsSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'
import {
  refreshExternalLinkCacheXml,
  type ImportedExternalLinkCaches,
  type ImportedExternalWorkbookReferences,
} from './xlsx-external-references.js'

const binaryChunkSize = 0x8000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const contentTypesPath = '[Content_Types].xml'
const externalLinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const externalLinkContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
const externalReferencesElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?externalReferences\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?externalReferences>)/u
const externalLinkPartPathPattern = /^xl\/externalLinks\/externalLink[1-9][0-9]*\.xml$/u
const externalLinkRelationshipPartPathPattern = /^xl\/externalLinks\/_rels\/externalLink[1-9][0-9]*\.xml\.rels$/u

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

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function extensionFromPath(path: string): string | null {
  const normalized = normalizeZipPath(path)
  const fileName = normalized.slice(normalized.lastIndexOf('/') + 1)
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

function isExternalLinkRelationship(relationship: ParsedRelationship | WorkbookPackageRelationshipSnapshot): boolean {
  return relationship.type === externalLinkRelationshipType
}

function externalLinkRelationshipPartPath(partPath: string): string {
  const normalized = normalizeZipPath(partPath)
  const fileName = normalized.slice(normalized.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function isExternalLinkPreservedPartPath(path: string): boolean {
  return externalLinkPartPathPattern.test(normalizeZipPath(path)) || externalLinkRelationshipPartPathPattern.test(normalizeZipPath(path))
}

function readWorkbookExternalReferencesXml(workbookXml: string | null): string | undefined {
  return workbookXml?.match(externalReferencesElementPattern)?.[0]
}

function readExternalLinkPartPaths(zip: XlsxZipEntries, workbookRelationships: readonly WorkbookPackageRelationshipSnapshot[]): string[] {
  const linkPartPaths = new Set<string>()
  for (const relationship of workbookRelationships) {
    const targetPath = normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))
    if (externalLinkPartPathPattern.test(targetPath) && zip[targetPath]) {
      linkPartPaths.add(targetPath)
    }
  }
  for (const path of Object.keys(zip)) {
    if (externalLinkPartPathPattern.test(path)) {
      linkPartPaths.add(path)
    }
  }

  const partPaths = new Set<string>(linkPartPaths)
  for (const linkPartPath of linkPartPaths) {
    const relsPath = externalLinkRelationshipPartPath(linkPartPath)
    if (zip[relsPath]) {
      partPaths.add(relsPath)
    }
  }
  return [...partPaths].toSorted()
}

function preservedPartsByPath(parts: readonly WorkbookPreservedPackagePartSnapshot[]): Map<string, Uint8Array> {
  const output = new Map<string, Uint8Array>()
  for (const part of parts) {
    if (!isExternalLinkPreservedPartPath(part.path)) {
      continue
    }
    const bytes = decodedPartBytes(part)
    if (bytes) {
      output.set(normalizeZipPath(part.path), bytes)
    }
  }
  return output
}

function externalLinkBookIndicesByPackagePath(references: ImportedExternalWorkbookReferences): Map<string, number> {
  const bookIndicesByPackagePath = new Map<string, number>()
  for (const reference of references.values()) {
    if (reference.packagePath) {
      bookIndicesByPackagePath.set(normalizeZipPath(reference.packagePath), reference.bookIndex)
    }
  }
  return bookIndicesByPackagePath
}

function relationshipTargetExists(
  relationship: WorkbookPackageRelationshipSnapshot,
  partsByPath: ReadonlyMap<string, Uint8Array>,
  zip: XlsxZipEntries,
): boolean {
  const targetPath = normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))
  return externalLinkPartPathPattern.test(targetPath) && (partsByPath.has(targetPath) || Boolean(zip[targetPath]))
}

function addWorkbookExternalLinkRelationships(
  relationships: ParsedRelationship[],
  additions: readonly WorkbookPackageRelationshipSnapshot[] | undefined,
  partsByPath: ReadonlyMap<string, Uint8Array>,
  zip: XlsxZipEntries,
): { readonly changed: boolean; readonly idMap: Map<string, string> } {
  const idMap = new Map<string, string>()
  let changed = false
  for (const addition of additions ?? []) {
    if (!isExternalLinkRelationship(addition) || !relationshipTargetExists(addition, partsByPath, zip)) {
      continue
    }
    const existing = relationships.find((relationship) => relationship.type === addition.type && relationship.target === addition.target)
    if (existing) {
      idMap.set(addition.id, existing.id)
      continue
    }
    const idInUse = relationships.some((relationship) => relationship.id === addition.id)
    const nextId = addition.id.length > 0 && !idInUse ? addition.id : nextRelationshipId(relationships)
    relationships.push({ ...parsedRelationship(addition), id: nextId })
    idMap.set(addition.id, nextId)
    changed = true
  }
  return { changed, idMap }
}

function replaceRelationshipIds(xml: string, relationshipIds: ReadonlyMap<string, string>): string {
  if (relationshipIds.size === 0) {
    return xml
  }
  return xml.replace(/\br:id=(["'])([\s\S]*?)\1/gu, (match, quote: string, id: string) => {
    const nextId = relationshipIds.get(id)
    return nextId ? match.replace(`${quote}${id}${quote}`, `${quote}${nextId}${quote}`) : match
  })
}

function insertWorkbookExternalReferencesXml(
  workbookXml: string,
  externalReferencesXml: string,
  relationshipIds: ReadonlyMap<string, string>,
): string {
  const nextExternalReferencesXml = replaceRelationshipIds(externalReferencesXml, relationshipIds)
  const withoutExternalReferences = ensureRelationshipNamespace(workbookXml.replace(externalReferencesElementPattern, ''))
  if (withoutExternalReferences.includes('</sheets>')) {
    return withoutExternalReferences.replace('</sheets>', `</sheets>${nextExternalReferencesXml}`)
  }
  if (/<(?:[A-Za-z_][\w.-]*:)?definedNames\b/u.test(withoutExternalReferences)) {
    return withoutExternalReferences.replace(/<(?:[A-Za-z_][\w.-]*:)?definedNames\b/u, `${nextExternalReferencesXml}$&`)
  }
  if (/<(?:[A-Za-z_][\w.-]*:)?calcPr\b/u.test(withoutExternalReferences)) {
    return withoutExternalReferences.replace(/<(?:[A-Za-z_][\w.-]*:)?calcPr\b/u, `${nextExternalReferencesXml}$&`)
  }
  return withoutExternalReferences.replace('</workbook>', `${nextExternalReferencesXml}</workbook>`)
}

function addExternalLinkContentTypes(
  contentTypesXml: string,
  artifacts: WorkbookExternalLinkArtifactsSnapshot,
  copiedPartPaths: ReadonlySet<string>,
): string {
  let output = contentTypesXml
  const copiedExtensions = new Set(
    [...copiedPartPaths].map(extensionFromPath).filter((extension): extension is string => Boolean(extension)),
  )
  for (const defaultEntry of artifacts.contentTypeDefaults ?? []) {
    if (copiedExtensions.has(defaultEntry.extension)) {
      output = addContentTypeDefault(output, defaultEntry.extension, defaultEntry.contentType)
    }
  }
  for (const overrideEntry of artifacts.contentTypeOverrides ?? []) {
    const path = normalizeZipPath(overrideEntry.partName)
    if (copiedPartPaths.has(path)) {
      output = upsertContentTypeOverride(output, overrideEntry.partName, overrideEntry.contentType)
    }
  }
  for (const path of copiedPartPaths) {
    if (externalLinkPartPathPattern.test(path)) {
      output = upsertContentTypeOverride(output, `/${path}`, externalLinkContentType)
    }
  }
  return output
}

export function readImportedWorkbookExternalLinkArtifacts(source: XlsxZipSource): WorkbookExternalLinkArtifactsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const workbookExternalReferencesXml = readWorkbookExternalReferencesXml(getZipText(zip, workbookPath))
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
    .filter(isExternalLinkRelationship)
    .map(relationshipSnapshot)
  const partPaths = readExternalLinkPartPaths(zip, workbookRelationships)
  const parts = partPaths.flatMap((path) => {
    const bytes = zip[path]
    return bytes ? [encodedPartSnapshot(path, bytes)] : []
  })
  if (parts.length === 0 && !workbookExternalReferencesXml && workbookRelationships.length === 0) {
    return undefined
  }

  const contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const contentTypeDefaults = contentTypesXml ? readContentTypeDefaults(contentTypesXml, partPaths) : []
  const contentTypeOverrides = contentTypesXml ? readContentTypeOverrides(contentTypesXml, partPaths) : []
  return {
    parts,
    ...(workbookExternalReferencesXml ? { workbookExternalReferencesXml } : {}),
    ...(workbookRelationships.length > 0 ? { workbookRelationships } : {}),
    ...(contentTypeDefaults.length > 0 ? { contentTypeDefaults } : {}),
    ...(contentTypeOverrides.length > 0 ? { contentTypeOverrides } : {}),
  }
}

export function refreshImportedWorkbookExternalLinkArtifactCaches(
  artifacts: WorkbookExternalLinkArtifactsSnapshot | undefined,
  caches: ImportedExternalLinkCaches,
  refreshedBookIndices: ReadonlySet<number>,
  references: ImportedExternalWorkbookReferences,
): WorkbookExternalLinkArtifactsSnapshot | undefined {
  if (!artifacts || refreshedBookIndices.size === 0) {
    return artifacts
  }

  let changed = false
  const bookIndicesByPackagePath = externalLinkBookIndicesByPackagePath(references)
  const parts = artifacts.parts.map((part) => {
    const bookIndex = bookIndicesByPackagePath.get(normalizeZipPath(part.path))
    const cache = bookIndex === undefined || !refreshedBookIndices.has(bookIndex) ? undefined : caches.get(bookIndex)
    if (!cache || part.storage !== 'base64') {
      return part
    }
    const bytes = decodedPartBytes(part)
    if (!bytes) {
      return part
    }
    const xml = new TextDecoder().decode(bytes)
    const nextXml = refreshExternalLinkCacheXml(xml, cache)
    if (nextXml === xml) {
      return part
    }
    changed = true
    return encodedPartSnapshot(part.path, new TextEncoder().encode(nextXml))
  })

  return changed ? { ...artifacts, parts } : artifacts
}

export function addExportExternalLinkArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const artifacts = snapshot.workbook.metadata?.externalLinkArtifacts
  if (!artifacts || artifacts.parts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(artifacts.parts)
  if (partsByPath.size === 0) {
    return bytes
  }

  let changed = false
  const copiedPartPaths = new Set<string>()
  for (const [path, partBytes] of partsByPath) {
    zip[path] = partBytes
    copiedPartPaths.add(path)
    changed = true
  }

  const relationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const relationshipResult = addWorkbookExternalLinkRelationships(relationships, artifacts.workbookRelationships, partsByPath, zip)
  if (relationshipResult.changed) {
    setZipText(zip, workbookRelationshipsPath, buildRelationshipsXml(relationships))
    changed = true
  }

  const workbookXml = getZipText(zip, workbookPath)
  if (workbookXml && artifacts.workbookExternalReferencesXml) {
    setZipText(
      zip,
      workbookPath,
      insertWorkbookExternalReferencesXml(workbookXml, artifacts.workbookExternalReferencesXml, relationshipResult.idMap),
    )
    changed = true
  }

  const contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const nextContentTypesXml = contentTypesXml ? addExternalLinkContentTypes(contentTypesXml, artifacts, copiedPartPaths) : contentTypesXml
  if (nextContentTypesXml !== contentTypesXml) {
    setZipText(zip, contentTypesPath, nextContentTypesXml)
    changed = true
  }

  return changed ? zipSync(zip) : bytes
}
