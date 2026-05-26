import type * as Protocol from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import type { SheetGridAxisRemapScope } from './sheet-grid.js'
import { CellStore } from './cell-store.js'
import type { StructuralTransaction } from './engine/structural-transaction.js'
import type { EngineCounters } from './perf/engine-counters.js'
import { createWorkbookMetadataService, runWorkbookMetadataEffect } from './workbook-metadata-service.js'
import {
  createWorkbookMetadataRecord,
  type WorkbookAxisMetadataRecord,
  type WorkbookCalculationSettingsRecord,
  type WorkbookChartRecord,
  type WorkbookConditionalFormatRecord,
  type WorkbookSheetConditionalFormatArtifactsRecord,
  type WorkbookDataValidationRecord,
  type WorkbookDrawingArtifactsRecord,
  type WorkbookExternalLinkArtifactsRecord,
  type WorkbookCellNumberFormatRecord,
  type WorkbookCellStyleRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFormatRangeRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookHyperlinkRecord,
  type WorkbookImageRecord,
  type WorkbookMacroPayloadRecord,
  type WorkbookMergeRangeRecord,
  type WorkbookMetadataRecord,
  type WorkbookPivotRecord,
  type WorkbookProtectionRecord,
  type WorkbookPropertyRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetDrawingArtifactsRecord,
  type WorkbookSheetProtectionRecord,
  type WorkbookSheetTabColorRecord,
  type WorkbookShapeRecord,
  type WorkbookSortKeyRecord,
  type WorkbookSortRecord,
  type WorkbookSpillRecord,
  type WorkbookStyleRangeRecord,
  type WorkbookTableRecord,
  type WorkbookVolatileContextRecord,
  type WorkbookNoteRecord,
} from './workbook-metadata-types.js'
import {
  coalesceStyleRanges as coalesceWorkbookStyleRanges,
  getCellStyle as readCellStyle,
  getCellNumberFormat as readCellNumberFormat,
  getRangeFormatId as readRangeFormatId,
  getStyleId as readStyleId,
  internCellNumberFormat as internWorkbookCellNumberFormat,
  internCellStyle as internWorkbookCellStyle,
  listCellNumberFormats as listWorkbookCellNumberFormats,
  listCellStyles as listWorkbookCellStyles,
  listFormatRanges as listWorkbookFormatRanges,
  listStyleRanges as listWorkbookStyleRanges,
  setFormatRange as storeFormatRange,
  setFormatRanges as replaceFormatRanges,
  setStyleRange as storeStyleRange,
  setStyleRanges as replaceStyleRanges,
  upsertCellNumberFormat as storeCellNumberFormat,
  upsertCellStyle as storeCellStyle,
} from './workbook-style-format-store.js'
import { WORKBOOK_DEFAULT_FORMAT_ID, WORKBOOK_DEFAULT_STYLE_ID, ensureWorkbookDefaultStyleFormat } from './workbook-default-style-format.js'
import { createCellKeyIndexMap } from './workbook-cell-key-index.js'
import { WorkbookCellRecordStore, type EnsuredCell } from './workbook-cell-record-store.js'
import { WorkbookAxisEntryStore } from './workbook-axis-entry-store.js'
import { WorkbookAxisMetadataStore, type WorkbookAxisGeometryPatch } from './workbook-axis-metadata-store.js'
import { WorkbookColumnVersionStore } from './workbook-column-version-store.js'
import { WorkbookIdAllocator } from './workbook-id-allocator.js'
import type { SheetRecord } from './workbook-sheet-record.js'
import { WorkbookSheetRegistryStore } from './workbook-sheet-registry-store.js'
import { WorkbookStructuralCellStore } from './workbook-structural-cell-store.js'
import { WorkbookStructuralAxisOperations } from './workbook-structural-axis-operations.js'
import { hasStructuralMetadataForSheetRecord, hasWorkbookMetadataForSheetRename } from './workbook-store-metadata-presence.js'
import { WorkbookStoreCommentAccessors } from './workbook-store-comment-accessors.js'

export { makeCellKey, makeLogicalCellKey } from './workbook-cell-key-index.js'
export { normalizeDefinedName, normalizeWorkbookObjectName, imageKey, pivotKey, shapeKey } from './workbook-metadata-types.js'
export type * from './workbook-store-types.js'

export class WorkbookStore extends WorkbookStoreCommentAccessors {
  static readonly defaultStyleId = WORKBOOK_DEFAULT_STYLE_ID
  static readonly defaultFormatId = WORKBOOK_DEFAULT_FORMAT_ID
  readonly cellStore = new CellStore()
  readonly sheetsByName = new Map<string, SheetRecord>()
  readonly sheetsById = new Map<number, SheetRecord>()
  readonly cellKeyToIndex: Map<number, number>
  readonly cellFormats = new Map<number, string>()
  readonly cellStyles = new Map<string, WorkbookCellStyleRecord>()
  readonly styleKeys = new Map<string, string>()
  readonly cellNumberFormats = new Map<string, WorkbookCellNumberFormatRecord>()
  readonly numberFormatKeys = new Map<string, string>()
  readonly metadata: WorkbookMetadataRecord = createWorkbookMetadataRecord()
  private readonly idAllocator = new WorkbookIdAllocator()
  protected readonly metadataService = createWorkbookMetadataService(this.metadata)
  private readonly sheetRegistry: WorkbookSheetRegistryStore
  private readonly cellRecordStore: WorkbookCellRecordStore
  private readonly axisEntryStore: WorkbookAxisEntryStore
  private readonly axisMetadataStore: WorkbookAxisMetadataStore
  private readonly columnVersionStore: WorkbookColumnVersionStore
  private readonly structuralCellStore: WorkbookStructuralCellStore
  private readonly structuralAxisOperations: WorkbookStructuralAxisOperations
  workbookName: string

  constructor(
    workbookName = 'Workbook',
    private readonly counters?: EngineCounters,
  ) {
    super()
    this.workbookName = workbookName
    this.cellKeyToIndex = createCellKeyIndexMap((sheetId, row, col) => this.getCellIndexAt(sheetId, row, col))
    this.sheetRegistry = new WorkbookSheetRegistryStore({
      sheetsByName: this.sheetsByName,
      sheetsById: this.sheetsById,
      metadata: this.metadata,
      counters: this.counters,
      cellKeyToIndex: this.cellKeyToIndex,
      cellFormats: this.cellFormats,
      getCellPosition: (cellIndex) => this.getCellPosition(cellIndex),
      deleteSheetRecords: (sheetName, context) => {
        runWorkbookMetadataEffect(this.metadataService.deleteSheetRecords(sheetName, context))
      },
      reorderSheetRecords: (context) => {
        runWorkbookMetadataEffect(this.metadataService.reorderSheetRecords(context))
      },
      renameSheetRecords: (oldName, nextName) => {
        if (!hasWorkbookMetadataForSheetRename(this.metadata)) {
          return
        }
        runWorkbookMetadataEffect(this.metadataService.renameSheet(oldName, nextName))
      },
    })
    this.cellRecordStore = new WorkbookCellRecordStore({
      cellStore: this.cellStore,
      cellKeyToIndex: this.cellKeyToIndex,
      cellFormats: this.cellFormats,
      getSheet: (sheetName) => this.getSheet(sheetName),
      getOrCreateSheet: (sheetName) => this.getOrCreateSheet(sheetName),
      getSheetById: (sheetId) => this.getSheetById(sheetId),
      getSheetNameById: (sheetId) => this.getSheetNameById(sheetId),
      createLogicalAxisId: (axis) => this.createLogicalAxisId(axis),
      createLogicalAxisIds: (axis, count) => this.createLogicalAxisIds(axis, count),
    })
    this.axisEntryStore = new WorkbookAxisEntryStore({
      counters: this.counters,
      createAxisEntry: (axis) => this.idAllocator.createAxisEntry(axis),
    })
    this.axisMetadataStore = new WorkbookAxisMetadataStore({
      axisEntryStore: this.axisEntryStore,
      metadata: this.metadata,
      getSheet: (sheetName) => this.getSheet(sheetName),
      getOrCreateSheet: (sheetName) => this.getOrCreateSheet(sheetName),
    })
    this.columnVersionStore = new WorkbookColumnVersionStore({
      cellStore: this.cellStore,
      getSheetById: (sheetId) => this.getSheetById(sheetId),
    })
    this.structuralCellStore = new WorkbookStructuralCellStore({
      counters: this.counters,
      cellStore: this.cellStore,
      cellKeyToIndex: this.cellKeyToIndex,
      getSheet: (sheetName) => this.getSheet(sheetName),
      createLogicalAxisId: (axis) => this.createLogicalAxisId(axis),
    })
    this.structuralAxisOperations = new WorkbookStructuralAxisOperations({
      axisEntryStore: this.axisEntryStore,
      getOrCreateSheet: (sheetName) => this.getOrCreateSheet(sheetName),
      bumpSheetStructureVersion: (sheet) => this.bumpSheetStructureVersion(sheet),
    })
    this.cellStore.onSetValue = (index) => {
      this.notifyCellValueWritten(index)
    }
    ensureWorkbookDefaultStyleFormat(this)
  }

  hasStructuralMetadataForSheet(sheetName: string): boolean {
    return (
      hasStructuralMetadataForSheetRecord(this.metadata, sheetName, this.getSheet(sheetName)) ||
      [...this.sheetsByName.values()].some((sheet) => sheet.sparklines !== undefined)
    )
  }

  hasProtectionMetadataForSheet(sheetName: string): boolean {
    if (this.metadata.sheetProtections.has(sheetName)) {
      return true
    }
    for (const protection of this.metadata.rangeProtections.values()) {
      if (protection.range.sheetName === sheetName) {
        return true
      }
    }
    return false
  }

  createSheet(name: string, order = this.sheetsByName.size, id?: number): SheetRecord {
    return this.sheetRegistry.createSheet(name, order, id)
  }

  moveSheet(name: string, order: number): SheetRecord | undefined {
    return this.sheetRegistry.moveSheet(name, order)
  }

  deleteSheet(name: string): void {
    this.sheetRegistry.deleteSheet(name)
  }

  renameSheet(oldName: string, nextName: string): SheetRecord | undefined {
    return this.sheetRegistry.renameSheet(oldName, nextName)
  }

  renameSheetById(sheetId: number, trimmedName: string): SheetRecord | undefined {
    return this.sheetRegistry.renameSheetById(sheetId, trimmedName)
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.sheetRegistry.getSheet(name)
  }

  getSheetColumnVersion(sheetName: string, col: number): number {
    return this.sheetRegistry.getSheetColumnVersion(sheetName, col)
  }

  getSheetStructureVersion(sheetName: string): number {
    return this.sheetRegistry.getSheetStructureVersion(sheetName)
  }

  getSheetById(id: number): SheetRecord | undefined {
    return this.sheetRegistry.getSheetById(id)
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name)
  }

  ensureCell(sheetName: string, address: string): number {
    return this.cellRecordStore.ensureCell(sheetName, address)
  }

  ensureCellRecord(sheetName: string, address: string): EnsuredCell {
    return this.cellRecordStore.ensureCellRecord(sheetName, address)
  }

  ensureCellAt(sheetId: number, row: number, col: number): EnsuredCell {
    return this.cellRecordStore.ensureCellAt(sheetId, row, col)
  }

  attachAllocatedCell(sheetId: number, row: number, col: number, cellIndex: number): void {
    this.cellRecordStore.attachAllocatedCell(sheetId, row, col, cellIndex)
  }

  ensureLogicalAxisId(sheetId: number, axis: 'row' | 'column', index: number): string {
    return this.cellRecordStore.ensureLogicalAxisId(sheetId, axis, index)
  }

  createLogicalAxisIdEnsurer(sheetId: number, axis: 'row' | 'column'): (index: number) => string {
    return this.cellRecordStore.createLogicalAxisIdEnsurer(sheetId, axis)
  }

  createDenseLogicalAxisIds(sheetId: number, axis: 'row' | 'column', start: number, count: number): string[] {
    return this.cellRecordStore.createDenseLogicalAxisIds(sheetId, axis, start, count)
  }

  attachAllocatedCellWithLogicalAxisIds(sheetId: number, row: number, col: number, cellIndex: number, rowId: string, colId: string): void {
    this.cellRecordStore.attachAllocatedCellWithLogicalAxisIds(sheetId, row, col, cellIndex, rowId, colId)
  }

  withBatchedColumnVersionUpdates<T>(execute: () => T): T {
    return this.columnVersionStore.withBatchedColumnVersionUpdates(execute)
  }

  notifyCellValueWritten(cellIndex: number): void {
    this.columnVersionStore.notifyCellValueWritten(cellIndex)
  }

  notifyColumnsWritten(sheetId: number, columns: readonly number[] | Uint32Array): void {
    this.columnVersionStore.notifyColumnsWritten(sheetId, columns)
  }

  getCellIndex(sheetName: string, address: string): number | undefined {
    return this.cellRecordStore.getCellIndex(sheetName, address)
  }

  getCellIndexAt(sheetId: number, row: number, col: number): number | undefined {
    return this.cellRecordStore.getCellIndexAt(sheetId, row, col)
  }

  getFreshCellIndexAt(sheetId: number, row: number, col: number): number | undefined {
    return this.cellRecordStore.getFreshCellIndexAt(sheetId, row, col)
  }

  getSheetNameById(id: number): string {
    return this.sheetRegistry.getSheetNameById(id)
  }

  getAddress(index: number): string {
    return this.cellRecordStore.getAddress(index)
  }

  getQualifiedAddress(index: number): string {
    return this.cellRecordStore.getQualifiedAddress(index)
  }

  getCellPosition(index: number): { sheetId: number; row: number; col: number } | undefined {
    return this.cellRecordStore.getCellPosition(index)
  }

  getCellAxisIndex(index: number, axis: 'row' | 'column'): number | undefined {
    return this.cellRecordStore.getCellAxisIndex(index, axis)
  }

  detachCellIndex(index: number): boolean {
    return this.cellRecordStore.detachCellIndex(index)
  }

  pruneCellIfEmpty(index: number): boolean {
    return this.cellRecordStore.pruneCellIfEmpty(index)
  }

  setCellFormat(index: number, format: string | null | undefined): void {
    if (format === undefined || format === null || format === '') {
      this.cellFormats.delete(index)
      return
    }
    this.internCellNumberFormat(format)
    this.cellFormats.set(index, format)
  }

  getCellFormat(index: number): string | undefined {
    return this.cellFormats.get(index)
  }

  upsertCellStyle(style: Protocol.CellStyleRecord): WorkbookCellStyleRecord {
    return storeCellStyle(this, style, (id) => this.idAllocator.bumpStyleId(id))
  }

  internCellStyle(style: Omit<WorkbookCellStyleRecord, 'id'>): WorkbookCellStyleRecord {
    return internWorkbookCellStyle(this, style, WorkbookStore.defaultStyleId)
  }

  getCellStyle(id: string | undefined): WorkbookCellStyleRecord | undefined {
    return readCellStyle(this, id, WorkbookStore.defaultStyleId)
  }

  getCellStyleProtection(sheetName: string, row: number, col: number): Protocol.CellStyleRecord['protection'] | undefined {
    return this.getCellStyle(this.getStyleId(sheetName, row, col))?.protection
  }

  listCellStyles(): WorkbookCellStyleRecord[] {
    return listWorkbookCellStyles(this)
  }

  upsertCellNumberFormat(format: Protocol.CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return storeCellNumberFormat(this, format, (id) => this.idAllocator.bumpFormatId(id))
  }

  internCellNumberFormat(format: string | Protocol.CellNumberFormatRecord): WorkbookCellNumberFormatRecord {
    return internWorkbookCellNumberFormat(this, format, WorkbookStore.defaultFormatId)
  }

  getCellNumberFormat(id: string | undefined): WorkbookCellNumberFormatRecord | undefined {
    return readCellNumberFormat(this, id, WorkbookStore.defaultFormatId)
  }

  listCellNumberFormats(): WorkbookCellNumberFormatRecord[] {
    return listWorkbookCellNumberFormats(this)
  }

  setStyleRange(range: Protocol.CellRangeRef, styleId: string): WorkbookStyleRangeRecord {
    return storeStyleRange(this, this.getOrCreateSheet(range.sheetName), range, styleId, WorkbookStore.defaultStyleId)
  }

  coalesceStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    const sheet = this.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    return coalesceWorkbookStyleRanges(sheet)
  }

  listStyleRanges(sheetName: string): WorkbookStyleRangeRecord[] {
    return listWorkbookStyleRanges(this.getSheet(sheetName))
  }

  setStyleRanges(sheetName: string, ranges: readonly Protocol.SheetStyleRangeSnapshot[]): WorkbookStyleRangeRecord[] {
    return replaceStyleRanges(this, this.getOrCreateSheet(sheetName), ranges)
  }

  getStyleId(sheetName: string, row: number, col: number): string {
    return readStyleId(this.getSheet(sheetName), row, col, WorkbookStore.defaultStyleId)
  }

  setFormatRange(range: Protocol.CellRangeRef, formatId: string): WorkbookFormatRangeRecord {
    return storeFormatRange(this, this.getOrCreateSheet(range.sheetName), range, formatId, WorkbookStore.defaultFormatId)
  }

  listFormatRanges(sheetName: string): WorkbookFormatRangeRecord[] {
    return listWorkbookFormatRanges(this.getSheet(sheetName))
  }

  setFormatRanges(sheetName: string, ranges: readonly Protocol.SheetFormatRangeSnapshot[]): WorkbookFormatRangeRecord[] {
    return replaceFormatRanges(this, this.getOrCreateSheet(sheetName), ranges)
  }

  getRangeFormatId(sheetName: string, row: number, col: number): string {
    return readRangeFormatId(this.getSheet(sheetName), row, col, WorkbookStore.defaultFormatId)
  }

  setWorkbookProperty(key: string, value: Protocol.LiteralInput): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProperty(key, value))
  }

  getWorkbookProperty(key: string): WorkbookPropertyRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProperty(key))
  }

  listWorkbookProperties(): WorkbookPropertyRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listWorkbookProperties())
  }

  setWorkbookProtection(record: Protocol.WorkbookProtectionSnapshot): WorkbookProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setWorkbookProtection(record))
  }

  getWorkbookProtection(): WorkbookProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getWorkbookProtection())
  }

  setMacroPayload(record: Protocol.WorkbookMacroPayloadSnapshot): WorkbookMacroPayloadRecord {
    return runWorkbookMetadataEffect(this.metadataService.setMacroPayload(record))
  }

  listMacroPayloads(): WorkbookMacroPayloadRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listMacroPayloads())
  }

  setCalculationSettings(settings: Protocol.WorkbookCalculationSettingsSnapshot): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setCalculationSettings(settings))
  }

  getCalculationSettings(): WorkbookCalculationSettingsRecord {
    return runWorkbookMetadataEffect(this.metadataService.getCalculationSettings())
  }

  setVolatileContext(context: Protocol.WorkbookVolatileContextSnapshot): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.setVolatileContext(context))
  }

  getVolatileContext(): WorkbookVolatileContextRecord {
    return runWorkbookMetadataEffect(this.metadataService.getVolatileContext())
  }

  setDrawingArtifacts(artifacts: Protocol.WorkbookDrawingArtifactsSnapshot): WorkbookDrawingArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDrawingArtifacts(artifacts))
  }

  getDrawingArtifacts(): WorkbookDrawingArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDrawingArtifacts())
  }

  clearDrawingArtifacts(): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearDrawingArtifacts())
  }

  setExternalLinkArtifacts(artifacts: Protocol.WorkbookExternalLinkArtifactsSnapshot): WorkbookExternalLinkArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setExternalLinkArtifacts(artifacts))
  }

  getExternalLinkArtifacts(): WorkbookExternalLinkArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getExternalLinkArtifacts())
  }

  clearExternalLinkArtifacts(): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearExternalLinkArtifacts())
  }

  setDefinedName(name: string, value: Protocol.WorkbookDefinedNameValueSnapshot, scopeSheetName?: string): WorkbookDefinedNameRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDefinedName(name, value, scopeSheetName))
  }

  getDefinedName(name: string, scopeSheetName?: string): WorkbookDefinedNameRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDefinedName(name, scopeSheetName))
  }

  deleteDefinedName(name: string, scopeSheetName?: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDefinedName(name, scopeSheetName))
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDefinedNames())
  }

  setTable(record: Protocol.WorkbookTableSnapshot): WorkbookTableRecord {
    return runWorkbookMetadataEffect(this.metadataService.setTable(record))
  }

  getTable(name: string): WorkbookTableRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getTable(name))
  }

  deleteTable(name: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteTable(name))
  }

  hasTables(): boolean {
    return this.metadata.tables.size > 0
  }

  listTables(): WorkbookTableRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listTables())
  }

  setRowMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
    geometry?: WorkbookAxisGeometryPatch,
    filterHidden?: boolean | null,
  ): WorkbookAxisMetadataRecord | undefined {
    return this.axisMetadataStore.setRowMetadata(sheetName, start, count, size, hidden, geometry, filterHidden)
  }

  getRowMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    return this.axisMetadataStore.getRowMetadata(sheetName, start, count)
  }

  listRowMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.axisMetadataStore.listRowMetadata(sheetName)
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
    return this.axisMetadataStore.setColumnMetadata(sheetName, start, count, size, hidden, geometry, filterHidden)
  }

  getColumnMetadata(sheetName: string, start: number, count: number): WorkbookAxisMetadataRecord | undefined {
    return this.axisMetadataStore.getColumnMetadata(sheetName, start, count)
  }

  listColumnMetadata(sheetName: string): WorkbookAxisMetadataRecord[] {
    return this.axisMetadataStore.listColumnMetadata(sheetName)
  }

  setSheetFormatPr(sheetName: string, sheetFormatPr: Protocol.WorkbookSheetFormatPrSnapshot): void {
    this.getOrCreateSheet(sheetName).sheetFormatPr = structuredClone(sheetFormatPr)
  }

  getSheetFormatPr(sheetName: string): Protocol.WorkbookSheetFormatPrSnapshot | undefined {
    const sheetFormatPr = this.getSheet(sheetName)?.sheetFormatPr
    return sheetFormatPr ? structuredClone(sheetFormatPr) : undefined
  }

  setSheetVisibility(sheetName: string, visibility: Protocol.WorkbookSheetVisibilitySnapshot | undefined): void {
    const sheet = this.getOrCreateSheet(sheetName)
    if (visibility === undefined) {
      delete sheet.visibility
      return
    }
    sheet.visibility = visibility
  }

  getSheetVisibility(sheetName: string): Protocol.WorkbookSheetVisibilitySnapshot | undefined {
    return this.getSheet(sheetName)?.visibility
  }

  listRowAxisEntries(sheetName: string): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.listAxisEntries(this.getSheet(sheetName), 'row')
  }

  listColumnAxisEntries(sheetName: string): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.listAxisEntries(this.getSheet(sheetName), 'column')
  }

  snapshotRowAxisEntries(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.snapshotAxisEntriesInRange(this.getSheet(sheetName), 'row', start, count)
  }

  snapshotColumnAxisEntries(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.snapshotAxisEntriesInRange(this.getSheet(sheetName), 'column', start, count)
  }

  materializeRowAxisEntries(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.materializeAxisEntries(this.getOrCreateSheet(sheetName), 'row', start, count)
  }

  materializeColumnAxisEntries(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.axisEntryStore.materializeAxisEntries(this.getOrCreateSheet(sheetName), 'column', start, count)
  }

  private bumpProvidedAxisEntryIds(axis: 'row' | 'column', entries: readonly Protocol.WorkbookAxisEntrySnapshot[] | undefined): void {
    entries?.forEach((entry) => {
      this.idAllocator.bumpAxisId(axis, entry.id)
    })
  }

  insertRows(sheetName: string, start: number, count: number, entries?: readonly Protocol.WorkbookAxisEntrySnapshot[]): void {
    this.bumpProvidedAxisEntryIds('row', entries)
    this.structuralAxisOperations.insert('row', sheetName, start, count, entries)
  }

  deleteRows(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.structuralAxisOperations.delete('row', sheetName, start, count)
  }

  moveRows(sheetName: string, start: number, count: number, target: number): void {
    this.structuralAxisOperations.move('row', sheetName, start, count, target)
  }

  insertColumns(sheetName: string, start: number, count: number, entries?: readonly Protocol.WorkbookAxisEntrySnapshot[]): void {
    this.bumpProvidedAxisEntryIds('column', entries)
    this.structuralAxisOperations.insert('column', sheetName, start, count, entries)
  }

  deleteColumns(sheetName: string, start: number, count: number): Protocol.WorkbookAxisEntrySnapshot[] {
    return this.structuralAxisOperations.delete('column', sheetName, start, count)
  }

  moveColumns(sheetName: string, start: number, count: number, target: number): void {
    this.structuralAxisOperations.move('column', sheetName, start, count, target)
  }

  setFreezePane(
    sheetName: string,
    rows: number,
    cols: number,
    options?: Pick<Protocol.WorkbookFreezePaneSnapshot, 'topLeftCell' | 'activePane'>,
  ): WorkbookFreezePaneRecord {
    return runWorkbookMetadataEffect(this.metadataService.setFreezePane(sheetName, rows, cols, options))
  }

  getFreezePane(sheetName: string): WorkbookFreezePaneRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFreezePane(sheetName))
  }

  clearFreezePane(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearFreezePane(sheetName))
  }

  setSheetTabColor(sheetName: string, tabColor: Protocol.WorkbookSheetTabColorSnapshot): WorkbookSheetTabColorRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetTabColor(sheetName, tabColor))
  }

  getSheetTabColor(sheetName: string): WorkbookSheetTabColorRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetTabColor(sheetName))
  }

  clearSheetTabColor(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearSheetTabColor(sheetName))
  }

  setMergeRange(range: Protocol.CellRangeRef): WorkbookMergeRangeRecord {
    return runWorkbookMetadataEffect(this.metadataService.setMergeRange(range))
  }

  setMergeRanges(sheetName: string, ranges: readonly Protocol.CellRangeRef[]): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.setMergeRanges(sheetName, ranges))
  }

  getMergeRange(sheetName: string, address: string): WorkbookMergeRangeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getMergeRange(sheetName, address))
  }

  getMergeRangeByRange(range: Protocol.CellRangeRef): WorkbookMergeRangeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getMergeRangeByRange(range))
  }

  clearMergeRanges(range: Protocol.CellRangeRef): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.clearMergeRanges(range))
  }

  listMergeRanges(sheetName: string): WorkbookMergeRangeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listMergeRanges(sheetName))
  }

  setSheetProtection(record: Protocol.WorkbookSheetProtectionSnapshot): WorkbookSheetProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetProtection(record))
  }

  getSheetProtection(sheetName: string): WorkbookSheetProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetProtection(sheetName))
  }

  clearSheetProtection(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.clearSheetProtection(sheetName))
  }

  setFilter(sheetName: string, range: Protocol.WorkbookAutoFilterSnapshot): WorkbookFilterRecord {
    return runWorkbookMetadataEffect(this.metadataService.setFilter(sheetName, range))
  }

  getFilter(sheetName: string, range: Protocol.CellRangeRef): WorkbookFilterRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getFilter(sheetName, range))
  }

  deleteFilter(sheetName: string, range: Protocol.CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteFilter(sheetName, range))
  }

  listFilters(sheetName: string): WorkbookFilterRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listFilters(sheetName))
  }

  setSort(sheetName: string, range: Protocol.CellRangeRef, keys: readonly WorkbookSortKeyRecord[]): WorkbookSortRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSort(sheetName, range, keys))
  }

  getSort(sheetName: string, range: Protocol.CellRangeRef): WorkbookSortRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSort(sheetName, range))
  }

  deleteSort(sheetName: string, range: Protocol.CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSort(sheetName, range))
  }

  listSorts(sheetName: string): WorkbookSortRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSorts(sheetName))
  }

  setDataValidation(record: Protocol.WorkbookDataValidationSnapshot): WorkbookDataValidationRecord {
    return runWorkbookMetadataEffect(this.metadataService.setDataValidation(record))
  }

  getDataValidation(sheetName: string, range: Protocol.CellRangeRef): WorkbookDataValidationRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getDataValidation(sheetName, range))
  }

  deleteDataValidation(sheetName: string, range: Protocol.CellRangeRef): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteDataValidation(sheetName, range))
  }

  listDataValidations(sheetName: string): WorkbookDataValidationRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listDataValidations(sheetName))
  }

  setConditionalFormat(record: Protocol.WorkbookConditionalFormatSnapshot): WorkbookConditionalFormatRecord {
    return runWorkbookMetadataEffect(this.metadataService.setConditionalFormat(record))
  }

  getConditionalFormat(id: string): WorkbookConditionalFormatRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getConditionalFormat(id))
  }

  deleteConditionalFormat(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteConditionalFormat(id))
  }

  listConditionalFormats(sheetName: string): WorkbookConditionalFormatRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listConditionalFormats(sheetName))
  }

  setConditionalFormatArtifacts(
    sheetName: string,
    artifacts: Protocol.WorkbookSheetConditionalFormatArtifactsSnapshot,
  ): WorkbookSheetConditionalFormatArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setConditionalFormatArtifacts(sheetName, artifacts))
  }

  getConditionalFormatArtifacts(sheetName: string): WorkbookSheetConditionalFormatArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getConditionalFormatArtifacts(sheetName))
  }

  deleteConditionalFormatArtifacts(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteConditionalFormatArtifacts(sheetName))
  }

  setSheetDrawingArtifacts(
    sheetName: string,
    artifacts: Protocol.WorkbookSheetDrawingArtifactsSnapshot,
  ): WorkbookSheetDrawingArtifactsRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSheetDrawingArtifacts(sheetName, artifacts))
  }

  getSheetDrawingArtifacts(sheetName: string): WorkbookSheetDrawingArtifactsRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSheetDrawingArtifacts(sheetName))
  }

  deleteSheetDrawingArtifacts(sheetName: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSheetDrawingArtifacts(sheetName))
  }

  setRangeProtection(record: Protocol.WorkbookRangeProtectionSnapshot): WorkbookRangeProtectionRecord {
    return runWorkbookMetadataEffect(this.metadataService.setRangeProtection(record))
  }

  getRangeProtection(id: string): WorkbookRangeProtectionRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getRangeProtection(id))
  }

  deleteRangeProtection(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteRangeProtection(id))
  }

  listRangeProtections(sheetName: string): WorkbookRangeProtectionRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listRangeProtections(sheetName))
  }

  setNote(record: Protocol.WorkbookNoteSnapshot): WorkbookNoteRecord {
    return runWorkbookMetadataEffect(this.metadataService.setNote(record))
  }

  getNote(sheetName: string, address: string): WorkbookNoteRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getNote(sheetName, address))
  }

  deleteNote(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteNote(sheetName, address))
  }

  listNotes(sheetName: string): WorkbookNoteRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listNotes(sheetName))
  }

  setHyperlink(record: Protocol.WorkbookHyperlinkSnapshot): WorkbookHyperlinkRecord {
    return runWorkbookMetadataEffect(this.metadataService.setHyperlink(record))
  }

  getHyperlink(sheetName: string, address: string): WorkbookHyperlinkRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getHyperlink(sheetName, address))
  }

  deleteHyperlink(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteHyperlink(sheetName, address))
  }

  listHyperlinks(sheetName: string): WorkbookHyperlinkRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listHyperlinks(sheetName))
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    return runWorkbookMetadataEffect(this.metadataService.setSpill(sheetName, address, rows, cols))
  }

  getSpill(sheetName: string, address: string): WorkbookSpillRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getSpill(sheetName, address))
  }

  deleteSpill(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteSpill(sheetName, address))
  }

  listSpills(): WorkbookSpillRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listSpills())
  }

  setPivot(record: Protocol.WorkbookPivotSnapshot): WorkbookPivotRecord {
    return runWorkbookMetadataEffect(this.metadataService.setPivot(record))
  }

  getPivot(sheetName: string, address: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivot(sheetName, address))
  }

  getPivotByKey(key: string): WorkbookPivotRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getPivotByKey(key))
  }

  deletePivot(sheetName: string, address: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deletePivot(sheetName, address))
  }

  hasPivots(): boolean {
    return this.metadata.pivots.size > 0
  }

  listPivots(): WorkbookPivotRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listPivots())
  }

  setChart(record: Protocol.WorkbookChartSnapshot): WorkbookChartRecord {
    return runWorkbookMetadataEffect(this.metadataService.setChart(record))
  }

  getChart(id: string): WorkbookChartRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getChart(id))
  }

  deleteChart(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteChart(id))
  }

  listCharts(): WorkbookChartRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listCharts())
  }

  setImage(record: Protocol.WorkbookImageSnapshot): WorkbookImageRecord {
    return runWorkbookMetadataEffect(this.metadataService.setImage(record))
  }

  getImage(id: string): WorkbookImageRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getImage(id))
  }

  deleteImage(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteImage(id))
  }

  listImages(): WorkbookImageRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listImages())
  }

  setShape(record: Protocol.WorkbookShapeSnapshot): WorkbookShapeRecord {
    return runWorkbookMetadataEffect(this.metadataService.setShape(record))
  }

  getShape(id: string): WorkbookShapeRecord | undefined {
    return runWorkbookMetadataEffect(this.metadataService.getShape(id))
  }

  deleteShape(id: string): boolean {
    return runWorkbookMetadataEffect(this.metadataService.deleteShape(id))
  }

  listShapes(): WorkbookShapeRecord[] {
    return runWorkbookMetadataEffect(this.metadataService.listShapes())
  }

  remapSheetCells(
    sheetName: string,
    axis: 'row' | 'column',
    remapIndex: (index: number) => number | undefined,
    scope?: SheetGridAxisRemapScope,
  ): { changedCellIndices: number[]; removedCellIndices: number[] } {
    return this.structuralCellStore.remapSheetCells(sheetName, axis, remapIndex, scope)
  }

  planStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    return this.structuralCellStore.planStructuralAxisTransform(sheetName, transform)
  }

  applyPlannedStructuralTransaction(transaction: StructuralTransaction): StructuralTransaction | undefined {
    return this.structuralCellStore.applyPlannedStructuralTransaction(transaction)
  }

  applyStructuralAxisTransform(sheetName: string, transform: StructuralAxisTransform): StructuralTransaction | undefined {
    return this.structuralCellStore.applyStructuralAxisTransform(sheetName, transform)
  }

  reset(workbookName = 'Workbook'): void {
    this.workbookName = workbookName
    this.sheetRegistry.reset()
    this.cellKeyToIndex.clear()
    this.cellFormats.clear()
    this.cellStyles.clear()
    this.styleKeys.clear()
    this.cellNumberFormats.clear()
    this.numberFormatKeys.clear()
    runWorkbookMetadataEffect(this.metadataService.reset())
    this.idAllocator.reset()
    this.cellStore.reset()
    ensureWorkbookDefaultStyleFormat(this)
  }

  private bumpSheetStructureVersion(sheet: SheetRecord): void {
    sheet.structureVersion += 1
  }

  private createLogicalAxisId(axis: 'row' | 'column'): string {
    return this.idAllocator.createLogicalAxisId(axis)
  }

  private createLogicalAxisIds(axis: 'row' | 'column', count: number): readonly string[] {
    return this.idAllocator.createLogicalAxisIds(axis, count)
  }
}
