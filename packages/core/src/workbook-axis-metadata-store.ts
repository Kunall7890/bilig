import type { WorkbookAxisMetadataSnapshot } from '@bilig/protocol'
import type { SheetRecord } from './workbook-sheet-record.js'
import type { WorkbookAxisEntryStore } from './workbook-axis-entry-store.js'
import type { WorkbookAxisMetadataRecord, WorkbookMetadataRecord } from './workbook-metadata-types.js'

type WorkbookAxis = 'row' | 'column'
export type WorkbookAxisGeometryPatch = Omit<WorkbookAxisMetadataSnapshot, 'start' | 'count' | 'size' | 'hidden' | 'filterHidden'>

export class WorkbookAxisMetadataStore {
  constructor(
    private readonly options: {
      readonly axisEntryStore: WorkbookAxisEntryStore
      readonly metadata: WorkbookMetadataRecord
      readonly getSheet: (sheetName: string) => SheetRecord | undefined
      readonly getOrCreateSheet: (sheetName: string) => SheetRecord
    },
  ) {}

  setRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
    geometry?: WorkbookAxisGeometryPatch,
    filterHidden?: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata('row', sheetName, start, count, size, hidden, geometry, filterHidden)
  }

  getRowMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    return this.getAxisMetadata('row', sheetName, start, count)
  }

  listRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata('row', sheetName)
  }

  setColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
    geometry?: WorkbookAxisGeometryPatch,
    filterHidden?: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.setAxisMetadata('column', sheetName, start, count, size, hidden, geometry, filterHidden)
  }

  getColumnMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    return this.getAxisMetadata('column', sheetName, start, count)
  }

  listColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.listAxisMetadata('column', sheetName)
  }

  private setAxisMetadata(
    axis: WorkbookAxis,
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
    geometry?: WorkbookAxisGeometryPatch,
    filterHidden?: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.options.axisEntryStore.setAxisMetadata(
      this.options.getOrCreateSheet(sheetName),
      axis,
      this.metadataForAxis(axis),
      sheetName,
      start,
      count,
      size,
      hidden,
      geometry,
      filterHidden,
    )
  }

  private getAxisMetadata(axis: WorkbookAxis, sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    const sheet = this.options.getSheet(sheetName)
    return sheet ? this.options.axisEntryStore.getAxisMetadataRecord(sheet, axis, sheetName, start, count) : undefined
  }

  private listAxisMetadata(axis: WorkbookAxis, sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.options.axisEntryStore.listAxisMetadata(this.options.getSheet(sheetName), this.metadataForAxis(axis), sheetName, axis)
  }

  private metadataForAxis(axis: WorkbookAxis): Map<string, WorkbookAxisMetadataRecord> {
    return axis === 'row' ? this.options.metadata.rowMetadata : this.options.metadata.columnMetadata
  }
}
