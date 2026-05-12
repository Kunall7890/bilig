import { unzipSync, zipSync } from 'fflate'

import type { WorkbookPackageRelationshipSnapshot, WorkbookSnapshot, WorkbookThemeArtifactSnapshot } from '@bilig/protocol'
import {
  buildRelationshipsXml,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

const workbookPartPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const contentTypesPath = '[Content_Types].xml'
const packageContentTypesNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'

export const themeRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
export const themeContentType = 'application/vnd.openxmlformats-officedocument.theme+xml'

function relationshipSnapshot(relationship: ParsedRelationship): WorkbookPackageRelationshipSnapshot {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function themeContentTypeOverride(contentTypesXml: string | null, partName: string): string | undefined {
  if (!contentTypesXml) {
    return undefined
  }
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  for (const match of contentTypesXml.matchAll(overridePattern)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function relationshipTargetForThemePath(theme: WorkbookThemeArtifactSnapshot): string {
  const target = theme.relationship.target
  return target.length > 0 ? target : theme.path.replace(/^xl\//u, '')
}

function sourceRelationshipId(relationships: readonly ParsedRelationship[], theme: WorkbookThemeArtifactSnapshot): string {
  if (theme.relationship.id.length === 0) {
    return nextRelationshipId(relationships)
  }
  return relationships.some((relationship) => relationship.id === theme.relationship.id)
    ? nextRelationshipId(relationships)
    : theme.relationship.id
}

function upsertThemeRelationship(relationships: ParsedRelationship[], theme: WorkbookThemeArtifactSnapshot): void {
  const existingThemeIndex = relationships.findIndex((relationship) => relationship.type === themeRelationshipType)
  const existingTheme = existingThemeIndex >= 0 ? relationships[existingThemeIndex] : undefined
  const nextRelationship = {
    id: existingTheme?.id ?? sourceRelationshipId(relationships, theme),
    type: themeRelationshipType,
    target: relationshipTargetForThemePath(theme),
    ...(theme.relationship.targetMode ? { targetMode: theme.relationship.targetMode } : {}),
  }
  if (existingThemeIndex >= 0) {
    relationships[existingThemeIndex] = nextRelationship
    return
  }
  relationships.push(nextRelationship)
}

function buildContentTypesXml(overridePartName: string, contentType: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Types xmlns="${packageContentTypesNamespace}">`,
    `<Override PartName="${escapeXml(overridePartName)}" ContentType="${escapeXml(contentType)}"/>`,
    '</Types>',
  ].join('')
}

function upsertContentTypeOverride(contentTypesXml: string | null, partName: string, contentType: string): string {
  if (!contentTypesXml || !contentTypesXml.includes('</Types>')) {
    return buildContentTypesXml(partName, contentType)
  }
  const escapedPartName = escapeXml(partName)
  const escapedContentType = escapeXml(contentType)
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  let replaced = false
  const nextXml = contentTypesXml.replace(overridePattern, (match: string, attributes: string) => {
    if (readXmlAttribute(attributes, 'PartName') !== partName) {
      return match
    }
    replaced = true
    return `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/>`
  })
  return replaced
    ? nextXml
    : nextXml.replace('</Types>', `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/></Types>`)
}

export function readImportedWorkbookThemeArtifact(source?: XlsxZipSource): WorkbookThemeArtifactSnapshot | undefined {
  if (!source) {
    return undefined
  }
  const zip = readXlsxZipEntries(source)
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const themeRelationship = workbookRelationships.find((relationship) => relationship.type === themeRelationshipType)
  if (!themeRelationship) {
    return undefined
  }
  const themePath = normalizeZipPath(resolveTargetPath(workbookPartPath, themeRelationship.target))
  const themeXml = getZipText(zip, themePath)
  if (!themeXml) {
    return undefined
  }
  const partName = `/${themePath}`
  const contentType = themeContentTypeOverride(getZipText(zip, contentTypesPath), partName) ?? themeContentType
  return {
    path: themePath,
    xml: themeXml,
    relationship: relationshipSnapshot(themeRelationship),
    contentType,
  }
}

export function addExportThemeArtifactToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const theme = snapshot.workbook.metadata?.styleArtifacts?.theme
  if (!theme) {
    return bytes
  }
  const zip = unzipSync(bytes)
  const themePath = normalizeZipPath(theme.path)
  setZipText(zip, themePath, theme.xml)

  const relationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  upsertThemeRelationship(relationships, theme)
  setZipText(zip, workbookRelationshipsPath, buildRelationshipsXml(relationships))

  setZipText(
    zip,
    contentTypesPath,
    upsertContentTypeOverride(getZipText(zip, contentTypesPath), `/${themePath}`, theme.contentType ?? themeContentType),
  )
  return zipSync(zip)
}
