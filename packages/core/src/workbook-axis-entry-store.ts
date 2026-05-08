import type { WorkbookAxisEntrySnapshot, WorkbookAxisMetadataSnapshot } from '@bilig/protocol'
import { addEngineCounter, type EngineCounters } from './perf/engine-counters.js'
import {
  getAxisMetadataRecord,
  listAxisEntries,
  materializeAxisEntries,
  materializeAxisEntryRecords,
  moveAxisEntries,
  snapshotAxisEntriesInRange,
  spliceAxisEntries,
  syncAxisMetadataBucket,
} from './workbook-axis-records.js'
import type { SheetRecord } from './workbook-sheet-record.js'
import { axisMetadataKey } from './workbook-store-records.js'
import type { WorkbookAxisEntryRecord, WorkbookAxisMetadataRecord } from './workbook-metadata-types.js'

type WorkbookAxis = 'row' | 'column'
type AxisGeometryPatch = Omit<WorkbookAxisMetadataSnapshot, 'start' | 'count' | 'size' | 'hidden'>

function applyAxisGeometryPatch(entry: WorkbookAxisEntryRecord, patch: AxisGeometryPatch | undefined): void {
  if (patch?.xlsxWidth !== undefined) {
    entry.xlsxWidth = patch.xlsxWidth
  } else {
    delete entry.xlsxWidth
  }
  if (patch?.xlsxHeight !== undefined) {
    entry.xlsxHeight = patch.xlsxHeight
  } else {
    delete entry.xlsxHeight
  }
  if (patch?.customWidth !== undefined) {
    entry.customWidth = patch.customWidth
  } else {
    delete entry.customWidth
  }
  if (patch?.bestFit !== undefined) {
    entry.bestFit = patch.bestFit
  } else {
    delete entry.bestFit
  }
  if (patch?.outlineLevel !== undefined) {
    entry.outlineLevel = patch.outlineLevel
  } else {
    delete entry.outlineLevel
  }
  if (patch?.collapsed !== undefined) {
    entry.collapsed = patch.collapsed
  } else {
    delete entry.collapsed
  }
  if (patch?.customHeight !== undefined) {
    entry.customHeight = patch.customHeight
  } else {
    delete entry.customHeight
  }
  if (patch?.thickTop !== undefined) {
    entry.thickTop = patch.thickTop
  } else {
    delete entry.thickTop
  }
  if (patch?.thickBottom !== undefined) {
    entry.thickBottom = patch.thickBottom
  } else {
    delete entry.thickBottom
  }
}

export class WorkbookAxisEntryStore {
  constructor(
    private readonly options: {
      readonly counters: EngineCounters | undefined
      readonly createAxisEntry: (axis: WorkbookAxis) => WorkbookAxisEntryRecord
    },
  ) {}

  setAxisMetadata(
    sheet: SheetRecord,
    axis: WorkbookAxis,
    bucket: Map<string, WorkbookAxisMetadataRecord>,
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
    geometry?: AxisGeometryPatch,
  ): WorkbookAxisMetadataRecord | undefined {
    const entries = this.materializeAxisEntryRecords(sheet, axis, start, count)
    entries.forEach((entry) => {
      entry.size = size
      entry.hidden = hidden
      applyAxisGeometryPatch(entry, geometry)
    })
    this.syncAxisMetadataBucket(sheetName, sheet, axis, bucket)
    const record = this.getAxisMetadataRecord(sheet, axis, sheetName, start, count)
    if (!record) {
      bucket.delete(axisMetadataKey(sheetName, start, count))
    }
    return record
  }

  listAxisMetadata(
    sheet: SheetRecord | undefined,
    bucket: Map<string, WorkbookAxisMetadataRecord>,
    sheetName: string,
    axis: WorkbookAxis,
  ): WorkbookAxisMetadataRecord[] {
    if (!sheet) {
      return []
    }
    this.syncAxisMetadataBucket(sheetName, sheet, axis, bucket)
    return [...bucket.values()]
      .filter((record) => record.sheetName === sheetName)
      .toSorted((left, right) => left.start - right.start || left.count - right.count)
  }

  listAxisEntries(sheet: SheetRecord | undefined, axis: WorkbookAxis): WorkbookAxisEntrySnapshot[] {
    if (!sheet) {
      return []
    }
    return listAxisEntries(axis === 'row' ? sheet.rowAxis : sheet.columnAxis)
  }

  materializeAxisEntries(sheet: SheetRecord, axis: WorkbookAxis, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    this.hydrateAxisEntriesFromMap(sheet, axis, start, count)
    const entries = materializeAxisEntries(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, count, () =>
      this.options.createAxisEntry(axis),
    )
    sheet.axisMap.replaceRange(axis, start, entries)
    return entries
  }

  snapshotAxisEntriesInRange(
    sheet: SheetRecord | undefined,
    axis: WorkbookAxis,
    start: number,
    count: number,
  ): WorkbookAxisEntrySnapshot[] {
    if (!sheet) {
      return []
    }
    return snapshotAxisEntriesInRange(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, count)
  }

  spliceAxisEntries(
    sheet: SheetRecord,
    axis: WorkbookAxis,
    start: number,
    deleteCount: number,
    insertCount: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): WorkbookAxisEntrySnapshot[] {
    const removed = spliceAxisEntries(
      axis === 'row' ? sheet.rowAxis : sheet.columnAxis,
      start,
      deleteCount,
      insertCount,
      () => this.options.createAxisEntry(axis),
      entries,
    )
    sheet.axisMap.splice(
      axis,
      start,
      deleteCount,
      insertCount,
      snapshotAxisEntriesInRange(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, insertCount),
    )
    sheet.logicalAxisMap.splice(axis, start, deleteCount, insertCount, [])
    if (this.options.counters) {
      addEngineCounter(this.options.counters, 'axisMapSplices')
    }
    return removed
  }

  moveAxisEntries(sheet: SheetRecord, axis: WorkbookAxis, start: number, count: number, target: number): void {
    sheet.axisMap.move(axis, start, count, target)
    sheet.logicalAxisMap.move(axis, start, count, target)
    if (this.options.counters) {
      addEngineCounter(this.options.counters, 'axisMapMoves')
    }
    moveAxisEntries(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, count, target, () => this.options.createAxisEntry(axis))
  }

  getAxisMetadataRecord(
    sheet: SheetRecord,
    axis: WorkbookAxis,
    sheetName: string,
    start: number,
    count: number,
  ): WorkbookAxisMetadataRecord | undefined {
    return getAxisMetadataRecord(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, sheetName, start, count)
  }

  private hydrateAxisEntriesFromMap(sheet: SheetRecord, axis: WorkbookAxis, start: number, count: number): void {
    if (count <= 0) {
      return
    }
    const entries = axis === 'row' ? sheet.rowAxis : sheet.columnAxis
    const snapshots = sheet.axisMap.snapshot(axis, start, count)
    for (let index = 0; index < snapshots.length; index += 1) {
      const snapshot = snapshots[index]!
      if (entries[snapshot.index]) {
        continue
      }
      entries[snapshot.index] = {
        id: snapshot.id,
        size: null,
        hidden: null,
      }
    }
  }

  private materializeAxisEntryRecords(sheet: SheetRecord, axis: WorkbookAxis, start: number, count: number): WorkbookAxisEntryRecord[] {
    this.hydrateAxisEntriesFromMap(sheet, axis, start, count)
    const entries = materializeAxisEntryRecords(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, count, () =>
      this.options.createAxisEntry(axis),
    )
    sheet.axisMap.replaceRange(axis, start, snapshotAxisEntriesInRange(axis === 'row' ? sheet.rowAxis : sheet.columnAxis, start, count))
    return entries
  }

  private syncAxisMetadataBucket(
    sheetName: string,
    sheet: SheetRecord,
    axis: WorkbookAxis,
    bucket: Map<string, WorkbookAxisMetadataRecord>,
  ): void {
    syncAxisMetadataBucket(bucket, sheetName, axis === 'row' ? sheet.rowAxis : sheet.columnAxis)
  }
}
