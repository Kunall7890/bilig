import { strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookPrinterSettingsSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

const binaryChunkSize = 0x8000
const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const printerSettingsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
const printerSettingsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings'
const printerSettingsPathPattern = /^xl\/printerSettings\/[^/]+\.bin$/u

const pageSetupTailElements = [
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
] as const

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

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const id = readAttribute(attributes, 'Id')
    const target = readAttribute(attributes, 'Target')
    const type = readAttribute(attributes, 'Type')
    return id && target && type ? [{ id, target, type }] : []
  })
}

function nextRelationshipId(relationships: readonly ParsedRelationship[]): string {
  let next = 1
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/u.exec(relationship.id)
    if (match) {
      next = Math.max(next, Number(match[1]) + 1)
    }
  }
  return `rId${String(next)}`
}

function relationshipXml(relationship: ParsedRelationship): string {
  return `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(relationship.target)}"/>`
}

function appendRelationshipXml(xml: string | null, relationship: ParsedRelationship): string {
  if (!xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">${relationshipXml(
      relationship,
    )}</Relationships>`
  }
  if (!xml.includes('</Relationships>')) {
    return xml
  }
  return xml.replace('</Relationships>', `${relationshipXml(relationship)}</Relationships>`)
}

function resolveTargetPath(basePartPath: string, target: string): string {
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

function readSheetRelationshipPath(sheetIndex: number): string {
  return `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
}

function readPageSetupXml(sheetXml: string | null, relationshipId: string): string | undefined {
  if (!sheetXml) {
    return undefined
  }
  const pageSetupXml = /<pageSetup\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetup>)/u.exec(sheetXml)?.[0]
  if (!pageSetupXml) {
    return undefined
  }
  const pageSetupRelationshipId = readAttribute(pageSetupXml, 'r:id') ?? readAttribute(pageSetupXml, 'id')
  return pageSetupRelationshipId === relationshipId ? pageSetupXml : undefined
}

export function readImportedWorkbookPrinterSettings(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookPrinterSettingsSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const printerSettingsBySheet = new Map<string, WorkbookPrinterSettingsSnapshot[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    const relationships = parseRelationships(getZipText(zip, readSheetRelationshipPath(sheetIndex))).filter(
      (entry) => entry.type === printerSettingsRelationshipType,
    )
    const settings = relationships.flatMap((relationship): WorkbookPrinterSettingsSnapshot[] => {
      const partPath = resolveTargetPath(sheetPath, relationship.target)
      const bytes = zip[normalizeZipPath(partPath)]
      if (!bytes) {
        return []
      }
      const pageSetupXml = readPageSetupXml(sheetXml, relationship.id)
      return [
        {
          relationshipTarget: relationship.target,
          storage: 'base64',
          dataBase64: encodeBase64(bytes),
          byteLength: bytes.byteLength,
          ...(pageSetupXml ? { pageSetupXml } : {}),
        },
      ]
    })
    if (settings.length > 0) {
      printerSettingsBySheet.set(sheetName, settings)
    }
  })

  return printerSettingsBySheet
}

function nextPartIndex(zip: XlsxZipEntries): number {
  let next = 1
  for (const path of Object.keys(zip)) {
    if (!path.startsWith('xl/printerSettings/printerSettings') || !path.endsWith('.bin')) {
      continue
    }
    const raw = path.slice('xl/printerSettings/printerSettings'.length, -'.bin'.length)
    const value = Number(raw)
    if (Number.isInteger(value) && value >= next) {
      next = value + 1
    }
  }
  return next
}

function safePrinterSettingsTarget(input: { readonly zip: XlsxZipEntries; readonly sheetPath: string; readonly target: string }): {
  readonly target: string
  readonly partPath: string
} {
  const partPath = normalizeZipPath(resolveTargetPath(input.sheetPath, input.target))
  if (printerSettingsPathPattern.test(partPath)) {
    return { target: input.target, partPath }
  }
  const index = nextPartIndex(input.zip)
  return {
    target: `../printerSettings/printerSettings${String(index)}.bin`,
    partPath: `xl/printerSettings/printerSettings${String(index)}.bin`,
  }
}

function addPrinterSettingsContentTypeOverride(contentTypesXml: string | null, partPath: string): string | null {
  if (!contentTypesXml) {
    return null
  }
  const partName = `/${normalizeZipPath(partPath)}`
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  if (!contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${escapeXml(partName)}" ContentType="${printerSettingsContentType}"/></Types>`,
  )
}

function ensureOfficeRelationshipNamespace(sheetXml: string): string {
  if (/\sxmlns:r=(["'])[\s\S]*?\1/u.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function setXmlAttribute(tag: string, name: string, value: string): string {
  const attribute = `${name}="${escapeXml(value)}"`
  const existingAttribute = new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u')
  if (existingAttribute.test(tag)) {
    return tag.replace(existingAttribute, ` ${attribute}`)
  }
  return tag.replace(/\/?>$/u, (ending) => ` ${attribute}${ending}`)
}

function pageSetupXmlForRelationship(pageSetupXml: string | undefined, relationshipId: string): string {
  const source = pageSetupXml ?? '<pageSetup/>'
  return source.replace(/<pageSetup\b[^>]*(?:\/>|>)/u, (opening) => setXmlAttribute(opening, 'r:id', relationshipId))
}

function insertOrReplacePageSetup(sheetXml: string, pageSetupXml: string): string {
  const withNamespace = ensureOfficeRelationshipNamespace(sheetXml)
  if (/<pageSetup\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetup>)/u.test(withNamespace)) {
    return withNamespace.replace(/<pageSetup\b[^>]*(?:\/>|>[\s\S]*?<\/pageSetup>)/u, pageSetupXml)
  }

  let insertIndex = withNamespace.indexOf('</worksheet>')
  for (const elementName of pageSetupTailElements) {
    const elementIndex = withNamespace.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return withNamespace
  }
  return `${withNamespace.slice(0, insertIndex)}${pageSetupXml}${withNamespace.slice(insertIndex)}`
}

function restorePrinterSettingsForSheet(input: {
  readonly zip: XlsxZipEntries
  readonly sheetIndex: number
  readonly setting: WorkbookPrinterSettingsSnapshot
}): boolean {
  if (input.setting.storage !== 'base64') {
    return false
  }
  const bytes = decodeBase64(input.setting.dataBase64)
  if (bytes.byteLength !== input.setting.byteLength) {
    return false
  }

  const sheetPath = `xl/worksheets/sheet${String(input.sheetIndex + 1)}.xml`
  const sheetXml = getZipText(input.zip, sheetPath)
  if (!sheetXml) {
    return false
  }
  const target = safePrinterSettingsTarget({ zip: input.zip, sheetPath, target: input.setting.relationshipTarget })
  input.zip[target.partPath] = bytes

  const sheetRelsPath = readSheetRelationshipPath(input.sheetIndex)
  const sheetRelsXml = getZipText(input.zip, sheetRelsPath)
  const relationships = parseRelationships(sheetRelsXml)
  const existingRelationship = relationships.find(
    (relationship) => relationship.type === printerSettingsRelationshipType && relationship.target === target.target,
  )
  const relationshipId = existingRelationship?.id ?? nextRelationshipId(relationships)
  if (!existingRelationship) {
    setZipText(
      input.zip,
      sheetRelsPath,
      appendRelationshipXml(sheetRelsXml, {
        id: relationshipId,
        target: target.target,
        type: printerSettingsRelationshipType,
      }),
    )
  }

  const pageSetupXml = pageSetupXmlForRelationship(input.setting.pageSetupXml, relationshipId)
  setZipText(input.zip, sheetPath, insertOrReplacePageSetup(sheetXml, pageSetupXml))

  const contentTypesXml = addPrinterSettingsContentTypeOverride(getZipText(input.zip, '[Content_Types].xml'), target.partPath)
  if (contentTypesXml) {
    setZipText(input.zip, '[Content_Types].xml', contentTypesXml)
  }
  return true
}

export function addExportPrinterSettingsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.printerSettings?.length ?? 0) > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      for (const setting of sheet.metadata?.printerSettings ?? []) {
        changed = restorePrinterSettingsForSheet({ zip, sheetIndex, setting }) || changed
      }
    })

  return changed ? zipSync(zip) : bytes
}
