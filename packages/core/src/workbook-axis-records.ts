import type { WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type { WorkbookAxisEntryRecord, WorkbookAxisMetadataRecord } from './workbook-metadata-types.js'
import { axisMetadataKey, deleteRecordsBySheet } from './workbook-store-records.js'

type AxisGeometryKey =
  | 'xlsxWidth'
  | 'xlsxHeight'
  | 'customWidth'
  | 'bestFit'
  | 'outlineLevel'
  | 'collapsed'
  | 'customHeight'
  | 'thickTop'
  | 'thickBottom'

const axisGeometryKeys = [
  'xlsxWidth',
  'xlsxHeight',
  'customWidth',
  'bestFit',
  'outlineLevel',
  'collapsed',
  'customHeight',
  'thickTop',
  'thickBottom',
] as const satisfies readonly AxisGeometryKey[]

type AxisGeometryMap = Partial<Record<AxisGeometryKey, number | boolean | null>>

function axisGeometryValue(entry: WorkbookAxisEntryRecord | WorkbookAxisEntrySnapshot, key: AxisGeometryKey): number | boolean | null {
  return entry[key] ?? null
}

function hasAxisGeometry(entry: WorkbookAxisEntryRecord): boolean {
  return axisGeometryKeys.some((key) => axisGeometryValue(entry, key) !== null)
}

function axisEntriesHaveSameMetadata(left: WorkbookAxisEntryRecord, right: WorkbookAxisEntryRecord): boolean {
  return (
    left.size === right.size &&
    left.hidden === right.hidden &&
    axisGeometryKeys.every((key) => axisGeometryValue(left, key) === axisGeometryValue(right, key))
  )
}

function copyAxisGeometry(source: WorkbookAxisEntryRecord | WorkbookAxisEntrySnapshot): Partial<WorkbookAxisEntrySnapshot> {
  return {
    ...(source.xlsxWidth !== undefined ? { xlsxWidth: source.xlsxWidth } : {}),
    ...(source.xlsxHeight !== undefined ? { xlsxHeight: source.xlsxHeight } : {}),
    ...(source.customWidth !== undefined ? { customWidth: source.customWidth } : {}),
    ...(source.bestFit !== undefined ? { bestFit: source.bestFit } : {}),
    ...(source.outlineLevel !== undefined ? { outlineLevel: source.outlineLevel } : {}),
    ...(source.collapsed !== undefined ? { collapsed: source.collapsed } : {}),
    ...(source.customHeight !== undefined ? { customHeight: source.customHeight } : {}),
    ...(source.thickTop !== undefined ? { thickTop: source.thickTop } : {}),
    ...(source.thickBottom !== undefined ? { thickBottom: source.thickBottom } : {}),
  }
}

function makeAxisEntrySnapshot(entry: WorkbookAxisEntryRecord, index: number): WorkbookAxisEntrySnapshot {
  const snapshot: WorkbookAxisEntrySnapshot = { id: entry.id, index, ...copyAxisGeometry(entry) }
  if (entry.size !== null) {
    snapshot.size = entry.size
  }
  if (entry.hidden !== null) {
    snapshot.hidden = entry.hidden
  }
  return snapshot
}

function makeAxisEntryRecord(snapshot: WorkbookAxisEntrySnapshot): WorkbookAxisEntryRecord {
  return {
    id: snapshot.id,
    size: snapshot.size ?? null,
    hidden: snapshot.hidden ?? null,
    ...copyAxisGeometry(snapshot),
  }
}

function copyAxisGeometryToMetadata(source: AxisGeometryMap): Partial<WorkbookAxisMetadataRecord> {
  const xlsxWidth = typeof source.xlsxWidth === 'number' ? source.xlsxWidth : undefined
  const xlsxHeight = typeof source.xlsxHeight === 'number' ? source.xlsxHeight : undefined
  const customWidth = typeof source.customWidth === 'boolean' ? source.customWidth : undefined
  const bestFit = typeof source.bestFit === 'boolean' ? source.bestFit : undefined
  const outlineLevel = typeof source.outlineLevel === 'number' ? source.outlineLevel : undefined
  const collapsed = typeof source.collapsed === 'boolean' ? source.collapsed : undefined
  const customHeight = typeof source.customHeight === 'boolean' ? source.customHeight : undefined
  const thickTop = typeof source.thickTop === 'boolean' ? source.thickTop : undefined
  const thickBottom = typeof source.thickBottom === 'boolean' ? source.thickBottom : undefined
  return {
    ...(xlsxWidth !== undefined ? { xlsxWidth } : {}),
    ...(xlsxHeight !== undefined ? { xlsxHeight } : {}),
    ...(customWidth !== undefined ? { customWidth } : {}),
    ...(bestFit !== undefined ? { bestFit } : {}),
    ...(outlineLevel !== undefined ? { outlineLevel } : {}),
    ...(collapsed !== undefined ? { collapsed } : {}),
    ...(customHeight !== undefined ? { customHeight } : {}),
    ...(thickTop !== undefined ? { thickTop } : {}),
    ...(thickBottom !== undefined ? { thickBottom } : {}),
  }
}

function makeAxisMetadataRecord(
  sheetName: string,
  start: number,
  count: number,
  entry: WorkbookAxisEntryRecord,
): WorkbookAxisMetadataRecord {
  return {
    sheetName,
    start,
    count,
    size: entry.size,
    hidden: entry.hidden,
    ...copyAxisGeometryToMetadata(
      Object.fromEntries(axisGeometryKeys.map((key) => [key, axisGeometryValue(entry, key)])) as AxisGeometryMap,
    ),
  }
}

export function listAxisEntries(entries: Array<WorkbookAxisEntryRecord | undefined>): WorkbookAxisEntrySnapshot[] {
  const result: WorkbookAxisEntrySnapshot[] = []
  entries.forEach((entry, index) => {
    if (!entry) {
      return
    }
    result.push(makeAxisEntrySnapshot(entry, index))
  })
  return result
}

export function materializeAxisEntryRecords(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  createEntry: () => WorkbookAxisEntryRecord,
): WorkbookAxisEntryRecord[] {
  const materialized: WorkbookAxisEntryRecord[] = []
  for (let index = 0; index < count; index += 1) {
    const position = start + index
    let entry = entries[position]
    if (!entry) {
      entry = createEntry()
      entries[position] = entry
    }
    materialized.push(entry)
  }
  return materialized
}

export function materializeAxisEntries(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  createEntry: () => WorkbookAxisEntryRecord,
): WorkbookAxisEntrySnapshot[] {
  return materializeAxisEntryRecords(entries, start, count, createEntry).map((entry, offset) =>
    makeAxisEntrySnapshot(entry, start + offset),
  )
}

export function snapshotAxisEntriesInRange(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
): WorkbookAxisEntrySnapshot[] {
  if (count <= 0) {
    return []
  }
  const snapshots: WorkbookAxisEntrySnapshot[] = []
  for (let offset = 0; offset < count; offset += 1) {
    const index = start + offset
    const entry = entries[index]
    if (!entry) {
      continue
    }
    snapshots.push(makeAxisEntrySnapshot(entry, index))
  }
  return snapshots
}

export function spliceAxisEntries(
  axisEntries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  deleteCount: number,
  insertCount: number,
  createEntry: () => WorkbookAxisEntryRecord,
  providedSnapshots?: readonly WorkbookAxisEntrySnapshot[],
): WorkbookAxisEntrySnapshot[] {
  const providedEntries = new Map<number, WorkbookAxisEntrySnapshot>()
  providedSnapshots?.forEach((entry) => {
    const offset = entry.index - start
    if (offset < 0 || offset >= insertCount) {
      return
    }
    providedEntries.set(offset, entry)
  })
  if (insertCount === 0 && axisEntries.length <= start) {
    return []
  }
  if (axisEntries.length < start) {
    axisEntries.length = start
  }
  const removed = axisEntries.splice(
    start,
    deleteCount,
    ...Array.from({ length: insertCount }, (_, index) => {
      const provided = providedEntries.get(index)
      if (provided) {
        return makeAxisEntryRecord(provided)
      }
      if (providedSnapshots) {
        return undefined
      }
      return createEntry()
    }),
  )
  return removed.flatMap((entry, index) => (entry ? [makeAxisEntrySnapshot(entry, start + index)] : []))
}

export function moveAxisEntries(
  axisEntries: Array<WorkbookAxisEntryRecord | undefined>,
  start: number,
  count: number,
  target: number,
  createEntry: () => WorkbookAxisEntryRecord,
): void {
  if (count <= 0 || start === target) {
    return
  }
  materializeAxisEntryRecords(axisEntries, start, count, createEntry)
  const moved = axisEntries.splice(start, count)
  axisEntries.splice(target, 0, ...moved)
}

export function getAxisMetadataRecord(
  entries: Array<WorkbookAxisEntryRecord | undefined>,
  sheetName: string,
  start: number,
  count: number,
): WorkbookAxisMetadataRecord | undefined {
  let size: number | null | undefined
  let hidden: boolean | null | undefined
  const geometry: AxisGeometryMap = {}
  let sawMaterialized = false
  for (let index = start; index < start + count; index += 1) {
    const entry = entries[index]
    if (!entry) {
      if (size === undefined) {
        size = null
      }
      if (hidden === undefined) {
        hidden = null
      }
      for (const key of axisGeometryKeys) {
        geometry[key] ??= null
      }
      continue
    }
    sawMaterialized = true
    size ??= entry.size
    hidden ??= entry.hidden
    if (size !== entry.size || hidden !== entry.hidden) {
      return undefined
    }
    for (const key of axisGeometryKeys) {
      const value = axisGeometryValue(entry, key)
      geometry[key] ??= value
      if (geometry[key] !== value) {
        return undefined
      }
    }
  }
  if (
    !sawMaterialized ||
    ((size ?? null) === null && (hidden ?? null) === null && axisGeometryKeys.every((key) => (geometry[key] ?? null) === null))
  ) {
    return undefined
  }
  return {
    sheetName,
    start,
    count,
    size: size ?? null,
    hidden: hidden ?? null,
    ...copyAxisGeometryToMetadata(geometry),
  }
}

export function syncAxisMetadataBucket(
  bucket: Map<string, WorkbookAxisMetadataRecord>,
  sheetName: string,
  entries: Array<WorkbookAxisEntryRecord | undefined>,
): void {
  deleteRecordsBySheet(bucket, sheetName, (record) => record.sheetName)
  let cursor = 0
  while (cursor < entries.length) {
    const entry = entries[cursor]
    if (!entry || (entry.size === null && entry.hidden === null && !hasAxisGeometry(entry))) {
      cursor += 1
      continue
    }
    const start = cursor
    cursor += 1
    while (cursor < entries.length) {
      const next = entries[cursor]
      if (!next || !axisEntriesHaveSameMetadata(entry, next)) {
        break
      }
      cursor += 1
    }
    const record = makeAxisMetadataRecord(sheetName, start, cursor - start, entry)
    bucket.set(axisMetadataKey(sheetName, start, record.count), record)
  }
}
