import type { SheetMetadataSnapshot, WorkbookMetadataSnapshot } from '@bilig/protocol'
import {
  preservedSheetMetadataKeys,
  preservedWorkbookMetadataKeys,
  type WorkbookPreservedMetadataRecord,
  type WorkbookPreservedSheetMetadataRecord,
} from './workbook-metadata-types.js'

function isCloneRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function clonePreservedMetadataValue(value: unknown): unknown {
  if (!isCloneRecord(value)) {
    return value
  }
  if (Array.isArray(value)) {
    return value.map(clonePreservedMetadataValue)
  }
  if (ArrayBuffer.isView(value)) {
    return structuredClone(value)
  }

  const path = value['path']
  const storage = value['storage']
  const dataBase64 = value['dataBase64']
  const byteLength = value['byteLength']
  if (typeof path === 'string' && storage === 'base64' && typeof dataBase64 === 'string' && typeof byteLength === 'number') {
    return { path, storage: 'base64', dataBase64, byteLength }
  }

  const xml = value['xml']
  if (typeof path === 'string' && typeof xml === 'string' && typeof value['readXml'] === 'function') {
    return { path, xml }
  }

  const cloned: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    const child = value[key]
    if (typeof child !== 'function') {
      cloned[key] = clonePreservedMetadataValue(child)
    }
  }
  return cloned
}

export function clonePreservedWorkbookMetadata(metadata: WorkbookPreservedMetadataRecord | undefined): WorkbookPreservedMetadataRecord {
  const cloned: WorkbookPreservedMetadataRecord = {}
  for (const key of preservedWorkbookMetadataKeys) {
    const value = metadata?.[key]
    if (value !== undefined) {
      Object.assign(cloned, { [key]: clonePreservedMetadataValue(value) })
    }
  }
  return cloned
}

export function clonePreservedSheetMetadata(
  metadata: WorkbookPreservedSheetMetadataRecord | undefined,
): WorkbookPreservedSheetMetadataRecord {
  const cloned: WorkbookPreservedSheetMetadataRecord = {}
  for (const key of preservedSheetMetadataKeys) {
    const value = metadata?.[key]
    if (value !== undefined) {
      Object.assign(cloned, { [key]: clonePreservedMetadataValue(value) })
    }
  }
  return cloned
}

export function pickPreservedWorkbookMetadata(metadata: WorkbookMetadataSnapshot | undefined): WorkbookPreservedMetadataRecord {
  return clonePreservedWorkbookMetadata(metadata)
}

export function pickPreservedSheetMetadata(metadata: SheetMetadataSnapshot | undefined): WorkbookPreservedSheetMetadataRecord {
  return clonePreservedSheetMetadata(metadata)
}

export function hasPreservedWorkbookMetadata(metadata: WorkbookPreservedMetadataRecord | undefined): boolean {
  return preservedWorkbookMetadataKeys.some((key) => metadata?.[key] !== undefined)
}

export function hasPreservedSheetMetadata(metadata: WorkbookPreservedSheetMetadataRecord | undefined): boolean {
  return preservedSheetMetadataKeys.some((key) => metadata?.[key] !== undefined)
}
