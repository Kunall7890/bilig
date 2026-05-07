import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { WorkbookCalculationSettingsSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const workbookCalcPrTailElements = [
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function insertWorkbookCalcPr(workbookXml: string, calcPrXml: string): string {
  if (/<calcPr\b/u.test(workbookXml)) {
    return workbookXml.replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/u, calcPrXml)
  }

  let insertIndex = workbookXml.indexOf('</workbook>')
  for (const elementName of workbookCalcPrTailElements) {
    const elementIndex = workbookXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return workbookXml
  }
  return `${workbookXml.slice(0, insertIndex)}${calcPrXml}${workbookXml.slice(insertIndex)}`
}

export function addExportCalculationSettingsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (snapshot.workbook.metadata?.calculationSettings?.mode !== 'manual') {
    return bytes
  }

  const zip = unzipSync(bytes)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return bytes
  }
  setZipText(zip, 'xl/workbook.xml', insertWorkbookCalcPr(workbookXml, '<calcPr calcMode="manual"/>'))
  return zipSync(zip)
}

export function readImportedWorkbookCalculationSettings(source: XlsxZipSource): WorkbookCalculationSettingsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return undefined
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const calcPr = recordChild(recordChild(parsed, 'workbook'), 'calcPr')
  if (calcPr?.['calcMode'] !== 'manual') {
    return undefined
  }
  return { mode: 'manual', compatibilityMode: 'excel-modern' }
}
