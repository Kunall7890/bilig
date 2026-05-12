import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookChartSheetArtifactsSnapshot,
  WorkbookDrawingArtifactsSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'

const binaryChunkSize = 0x8000
const chartSheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const chartSheetPartPathPattern = /^xl\/chartsheets\/sheet\d+\.xml$/u
const chartSheetRelsPathPattern = /^xl\/chartsheets\/_rels\/sheet\d+\.xml\.rels$/u

interface WorkbookSheetRelationshipEntry {
  readonly name: string
  readonly relationshipId: string
  readonly sheetId?: number
  readonly state?: 'hidden' | 'veryHidden'
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

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function decodeXmlAttribute(value: string): string {
  return value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&')
}

function workbookSheetRelationshipEntries(workbookXml: string | null): WorkbookSheetRelationshipEntry[] {
  if (!workbookXml) {
    return []
  }
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const rawName = readXmlAttribute(attributes, 'name')
    const relationshipId = readXmlAttribute(attributes, 'r:id') ?? readXmlAttribute(attributes, 'id')
    const sheetId = Number(readXmlAttribute(attributes, 'sheetId'))
    const state = readXmlAttribute(attributes, 'state')
    return rawName && relationshipId
      ? [
          {
            name: decodeXmlAttribute(rawName),
            relationshipId,
            ...(Number.isSafeInteger(sheetId) && sheetId > 0 ? { sheetId } : {}),
            ...(state === 'hidden' || state === 'veryHidden' ? { state } : {}),
          },
        ]
      : []
  })
}

function packageRelationshipsPath(partPath: string): string {
  const normalizedPath = normalizeZipPath(partPath)
  const directory = normalizedPath.slice(0, normalizedPath.lastIndexOf('/'))
  const fileName = normalizedPath.slice(normalizedPath.lastIndexOf('/') + 1)
  return `${directory}/_rels/${fileName}.rels`
}

function collectPackageDependencyPaths(zip: XlsxZipEntries, rootPartPath: string): Set<string> {
  const collectedPaths = new Set<string>()
  const pending = [normalizeZipPath(rootPartPath)]
  while (pending.length > 0) {
    const partPath = pending.pop()
    if (!partPath || collectedPaths.has(partPath) || !zip[partPath]) {
      continue
    }
    collectedPaths.add(partPath)
    const relsPath = packageRelationshipsPath(partPath)
    const relsXml = getZipText(zip, relsPath)
    if (!relsXml) {
      continue
    }
    collectedPaths.add(relsPath)
    parseRelationships(relsXml).forEach((relationship) => {
      if (relationship.targetMode === 'External') {
        return
      }
      pending.push(normalizeZipPath(resolveTargetPath(partPath, relationship.target)))
    })
  }
  return collectedPaths
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
    const extension = readXmlAttribute(attributes, 'Extension')?.toLowerCase()
    const contentType = readXmlAttribute(attributes, 'ContentType')
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
    const partName = readXmlAttribute(attributes, 'PartName')
    const contentType = readXmlAttribute(attributes, 'ContentType')
    if (!partName || !contentType || !neededPartNames.has(partName)) {
      continue
    }
    overridesByPartName.set(partName, { partName, contentType })
  }
  return [...overridesByPartName.values()].toSorted((left, right) => left.partName.localeCompare(right.partName))
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  const existingDefaultPattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${extension}\\1`, 'u')
  if (existingDefaultPattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`)
}

function removeContentTypeOverride(contentTypesXml: string, partName: string): string {
  return contentTypesXml.replace(new RegExp(`<Override\\b[^>]*\\bPartName="${partName}"[^>]*/>`, 'u'), '')
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeRelationshipTarget(target: string): string {
  return normalizeZipPath(target).replace(/^xl\//u, '')
}

function replaceWorkbookSheetRelationship(
  relationships: readonly ParsedRelationship[],
  relationshipId: string,
  target: string,
): ParsedRelationship[] {
  return relationships.map((relationship) =>
    relationship.id === relationshipId
      ? {
          id: relationship.id,
          type: chartSheetRelationshipType,
          target,
        }
      : relationship,
  )
}

function workbookSheetByName(workbookXml: string | null, sheetName: string): WorkbookSheetRelationshipEntry | undefined {
  return workbookSheetRelationshipEntries(workbookXml).find((sheet) => sheet.name === sheetName)
}

function nextWorkbookSheetId(workbookXml: string | null, preferredSheetId: number | undefined): number {
  const usedIds = new Set(workbookSheetRelationshipEntries(workbookXml).flatMap((sheet) => (sheet.sheetId ? [sheet.sheetId] : [])))
  if (preferredSheetId !== undefined && !usedIds.has(preferredSheetId)) {
    return preferredSheetId
  }
  let next = 1
  while (usedIds.has(next)) {
    next += 1
  }
  return next
}

function ensureWorkbookSheetEntry(workbookXml: string, entry: WorkbookChartSheetArtifactsSnapshot, relationshipId: string): string {
  if (workbookSheetByName(workbookXml, entry.name)) {
    return workbookXml
  }
  const sheetId = nextWorkbookSheetId(workbookXml, entry.sheetId)
  const stateAttribute = entry.state ? ` state="${entry.state}"` : ''
  const sheetXml = `<sheet name="${escapeXml(entry.name)}" sheetId="${String(sheetId)}" r:id="${escapeXml(relationshipId)}"${stateAttribute}/>`
  const withRelationshipNamespace = /xmlns:r=/u.test(workbookXml)
    ? workbookXml
    : workbookXml.replace(
        /<workbook\b([^>]*)>/u,
        `<workbook$1 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`,
      )
  return withRelationshipNamespace.includes('</sheets>')
    ? withRelationshipNamespace.replace('</sheets>', `${sheetXml}</sheets>`)
    : withRelationshipNamespace.replace('</workbook>', `<sheets>${sheetXml}</sheets></workbook>`)
}

export function readImportedWorkbookChartArtifacts(source: XlsxZipSource): {
  readonly artifacts: WorkbookDrawingArtifactsSnapshot | undefined
  readonly chartSheetArtifacts: WorkbookChartSheetArtifactsSnapshot[] | undefined
} {
  const zip = readXlsxZipEntries(source)
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const chartSheetArtifacts: WorkbookChartSheetArtifactsSnapshot[] = []
  const partPaths = new Set<string>()

  workbookSheetRelationshipEntries(getZipText(zip, 'xl/workbook.xml')).forEach((sheet) => {
    const relationship = workbookRelationships.find(
      (candidate) => candidate.id === sheet.relationshipId && candidate.type === chartSheetRelationshipType,
    )
    if (!relationship || relationship.targetMode === 'External') {
      return
    }
    const chartSheetPath = normalizeZipPath(resolveTargetPath('xl/workbook.xml', relationship.target))
    if (!chartSheetPartPathPattern.test(chartSheetPath)) {
      return
    }
    chartSheetArtifacts.push({
      name: sheet.name,
      relationshipTarget: relationship.target,
      ...(sheet.sheetId !== undefined ? { sheetId: sheet.sheetId } : {}),
      ...(sheet.state !== undefined ? { state: sheet.state } : {}),
    })
    collectPackageDependencyPaths(zip, chartSheetPath).forEach((path) => {
      partPaths.add(path)
    })
  })

  if (partPaths.size === 0) {
    return { artifacts: undefined, chartSheetArtifacts: undefined }
  }

  const contentTypePartPaths = [...partPaths].filter((path) => !path.endsWith('.rels'))
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
            ...(readContentTypeDefaults(contentTypesXml, contentTypePartPaths).length > 0
              ? { contentTypeDefaults: readContentTypeDefaults(contentTypesXml, contentTypePartPaths) }
              : {}),
            ...(readContentTypeOverrides(contentTypesXml, contentTypePartPaths).length > 0
              ? { contentTypeOverrides: readContentTypeOverrides(contentTypesXml, contentTypePartPaths) }
              : {}),
          }
        : {}),
    },
    chartSheetArtifacts: chartSheetArtifacts.length > 0 ? chartSheetArtifacts : undefined,
  }
}

export function addExportChartArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookArtifacts = snapshot.workbook.metadata?.chartArtifacts
  const chartSheetArtifacts = snapshot.workbook.metadata?.chartSheetArtifacts ?? []
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  if (!workbookArtifacts || chartSheetArtifacts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(workbookArtifacts.parts)
  const copiedPartPaths = new Set<string>()
  let changed = false

  partsByPath.forEach((partBytes, path) => {
    zip[path] = partBytes
    copiedPartPaths.add(path)
    changed = true
  })

  let workbookXml = getZipText(zip, 'xl/workbook.xml')
  let workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  chartSheetArtifacts.forEach((chartSheet) => {
    if (!workbookXml) {
      return
    }
    const existingSheet = workbookSheetByName(workbookXml, chartSheet.name)
    const relationshipId = existingSheet?.relationshipId ?? nextRelationshipId(workbookRelationships)
    const existingRelationship = workbookRelationships.find((relationship) => relationship.id === relationshipId)
    if (existingRelationship?.type === worksheetRelationshipType) {
      const generatedWorksheetPath = normalizeZipPath(resolveTargetPath('xl/workbook.xml', existingRelationship.target))
      delete zip[generatedWorksheetPath]
      delete zip[packageRelationshipsPath(generatedWorksheetPath)]
      changed = true
    }
    if (!existingRelationship) {
      workbookRelationships.push({
        id: relationshipId,
        type: chartSheetRelationshipType,
        target: chartSheet.relationshipTarget,
      })
    }
    workbookRelationships = replaceWorkbookSheetRelationship(workbookRelationships, relationshipId, chartSheet.relationshipTarget)
    workbookXml = ensureWorkbookSheetEntry(workbookXml, chartSheet, relationshipId)
    changed = true
  })
  setZipText(zip, 'xl/_rels/workbook.xml.rels', buildRelationshipsXml(workbookRelationships))
  if (workbookXml) {
    setZipText(zip, 'xl/workbook.xml', workbookXml)
  }

  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  chartSheetArtifacts.forEach((chartSheet) => {
    const sheetIndex = orderedSheets.findIndex((candidate) => candidate.name === chartSheet.name)
    if (sheetIndex >= 0) {
      contentTypesXml = removeContentTypeOverride(contentTypesXml, `/xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    }
  })
  for (const defaultEntry of workbookArtifacts.contentTypeDefaults ?? []) {
    contentTypesXml = addContentTypeDefault(contentTypesXml, defaultEntry.extension, defaultEntry.contentType)
  }
  for (const overrideEntry of workbookArtifacts.contentTypeOverrides ?? []) {
    if (!copiedPartPaths.has(normalizeZipPath(overrideEntry.partName))) {
      continue
    }
    contentTypesXml = addContentTypeOverride(contentTypesXml, overrideEntry.partName, overrideEntry.contentType)
  }
  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }

  const copiedChartSheetParts = [...copiedPartPaths].some(
    (path) => chartSheetPartPathPattern.test(path) || chartSheetRelsPathPattern.test(path),
  )
  const workbookHasChartSheetRelationship = workbookRelationships.some(
    (relationship) =>
      relationship.type === chartSheetRelationshipType &&
      copiedPartPaths.has(normalizeZipPath(resolveTargetPath('xl/workbook.xml', relationship.target))) &&
      normalizeRelationshipTarget(relationship.target).startsWith('chartsheets/'),
  )

  return changed && copiedChartSheetParts && workbookHasChartSheetRelationship ? zipSync(zip) : bytes
}
