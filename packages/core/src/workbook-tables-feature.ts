import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef, type CellValue, type WorkbookTableSnapshot } from '@bilig/protocol'
import {
  defineWorkbookFeaturePlugin,
  type WorkbookActionInput,
  type WorkbookCommandDescriptor,
  type WorkbookCommandReceipt,
  type WorkbookCommandRequest,
  type EngineOp,
  type WorkbookFeaturePlugin,
  type WorkbookRangeChromeProjection,
} from '@bilig/workbook'
import type { SpreadsheetEngine } from './engine.js'
import { buildDeleteTableOps, buildSetTableOps } from './engine/engine-workbook-object-metadata-ops.js'
import { excelCompatibleTableColumnName } from './engine/services/table-column-name-helpers.js'
import type { WorkbookCommandService } from './workbook-command-service.js'
import type { WorkbookProjectionInterceptorService } from './workbook-projection-interceptors.js'

export const WORKBOOK_TABLES_FEATURE_ID = 'tables'

export const WORKBOOK_TABLE_COMMAND_IDS = Object.freeze({
  createFromSelection: 'tables.createFromSelection',
  upsert: 'tables.upsert',
  delete: 'tables.delete',
  resize: 'tables.resize',
  renameHeader: 'tables.renameHeader',
})

const TABLE_COMMAND_DESCRIPTORS: readonly WorkbookCommandDescriptor[] = Object.freeze([
  {
    id: WORKBOOK_TABLE_COMMAND_IDS.createFromSelection,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    category: 'command',
    label: 'Create table from selection',
  },
  {
    id: WORKBOOK_TABLE_COMMAND_IDS.upsert,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    category: 'operation',
    label: 'Upsert table',
  },
  {
    id: WORKBOOK_TABLE_COMMAND_IDS.delete,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    category: 'operation',
    label: 'Delete table',
  },
  {
    id: WORKBOOK_TABLE_COMMAND_IDS.resize,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    category: 'operation',
    label: 'Resize table',
  },
  {
    id: WORKBOOK_TABLE_COMMAND_IDS.renameHeader,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    category: 'command',
    label: 'Rename table header',
  },
])

export function createWorkbookTablesFeaturePlugin(): WorkbookFeaturePlugin {
  return defineWorkbookFeaturePlugin({
    id: WORKBOOK_TABLES_FEATURE_ID,
    version: '1.0.0',
    commands: TABLE_COMMAND_DESCRIPTORS,
    projectionInterceptors: [
      {
        id: 'tables.rangeChrome',
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        point: 'rangeChrome',
        priority: 50,
        label: 'Table range chrome',
      },
      {
        id: 'tables.commandMetadata',
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        point: 'commandMetadata',
        priority: 50,
        label: 'Table command metadata',
      },
    ],
    uiContributions: [
      {
        id: 'tables.toolbar.create',
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        slot: 'toolbar',
        label: 'Create table',
        order: 40,
      },
      {
        id: 'tables.sidePanel.inspect',
        featureId: WORKBOOK_TABLES_FEATURE_ID,
        slot: 'sidePanel',
        label: 'Tables',
        order: 30,
      },
    ],
  })
}

export function registerWorkbookTablesFeature(args: {
  readonly commandService: WorkbookCommandService
  readonly projectionService: WorkbookProjectionInterceptorService
}): () => void {
  const disposers = [
    ...TABLE_COMMAND_DESCRIPTORS.map((descriptor) => args.commandService.registerCommand(descriptor, handleTableCommand)),
    args.projectionService.register({
      id: 'tables.rangeChrome',
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      point: 'rangeChrome',
      priority: 50,
      projectRangeChrome(input, context) {
        return projectTableRangeChrome(context.engine, input.range)
      },
    }),
    args.projectionService.register({
      id: 'tables.commandMetadata',
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      point: 'commandMetadata',
      priority: 50,
      projectCommandMetadata(input) {
        return tableCommandMetadata(input.request)
      },
    }),
  ]
  return () => {
    disposers.toReversed().forEach((dispose) => {
      dispose()
    })
  }
}

function handleTableCommand(request: WorkbookCommandRequest, context: { readonly engine: SpreadsheetEngine }): WorkbookCommandReceipt {
  switch (request.commandId) {
    case WORKBOOK_TABLE_COMMAND_IDS.createFromSelection:
      return executeCreateTableFromSelection(context.engine, request)
    case WORKBOOK_TABLE_COMMAND_IDS.upsert:
      return executeUpsertTable(context.engine, request)
    case WORKBOOK_TABLE_COMMAND_IDS.delete:
      return executeDeleteTable(context.engine, request)
    case WORKBOOK_TABLE_COMMAND_IDS.resize:
      return executeResizeTable(context.engine, request)
    case WORKBOOK_TABLE_COMMAND_IDS.renameHeader:
      return executeRenameHeader(context.engine, request)
    default:
      return rejectedReceipt(request, `Unknown tables command: ${request.commandId}`)
  }
}

function executeCreateTableFromSelection(engine: SpreadsheetEngine, request: WorkbookCommandRequest): WorkbookCommandReceipt {
  const input = inputRecord(request)
  const range = readRange(input['range'], 'range')
  const bounds = rangeBounds(range)
  const hasHeaders = typeof input['hasHeaders'] === 'boolean' ? input['hasHeaders'] : detectHeaderRow(engine, bounds)
  const name = typeof input['name'] === 'string' && input['name'].trim() !== '' ? input['name'].trim() : nextTableName(engine)
  const columnNames = tableColumnNamesForRange(engine, bounds, hasHeaders)
  const table: WorkbookTableSnapshot = {
    name,
    sheetName: range.sheetName,
    startAddress: formatAddress(bounds.startRow, bounds.startCol),
    endAddress: formatAddress(bounds.endRow, bounds.endCol),
    columnNames,
    columns: columnNames.map((columnName) => ({ name: columnName })),
    headerRow: hasHeaders,
    totalsRow: false,
  }
  return applyTableOps(engine, request, buildSetTableOps(engine.workbook, table), {
    message: `Create table ${table.name} on ${rangeLabel(tableRange(table))}`,
    proof: tableProof(table, 'create'),
    changedRanges: tableSemanticRanges(table),
  })
}

function executeUpsertTable(engine: SpreadsheetEngine, request: WorkbookCommandRequest): WorkbookCommandReceipt {
  const input = inputRecord(request)
  const table = readTable(input['table'])
  return applyTableOps(engine, request, buildSetTableOps(engine.workbook, table), {
    message: `Upsert table ${table.name}`,
    proof: tableProof(table, 'upsert'),
    changedRanges: tableSemanticRanges(table),
  })
}

function executeDeleteTable(engine: SpreadsheetEngine, request: WorkbookCommandRequest): WorkbookCommandReceipt {
  const input = inputRecord(request)
  const name = readNonEmptyString(input['name'], 'name')
  const existing = engine.getTable(name)
  if (!existing) {
    return noopReceipt(request, `Table ${name} does not exist`)
  }
  const ops = buildDeleteTableOps(engine.workbook, name) ?? []
  return applyTableOps(engine, request, ops, {
    message: `Delete table ${name}`,
    proof: tableProof(existing, 'delete'),
    changedRanges: tableSemanticRanges(existing),
  })
}

function executeResizeTable(engine: SpreadsheetEngine, request: WorkbookCommandRequest): WorkbookCommandReceipt {
  const input = inputRecord(request)
  const name = readNonEmptyString(input['name'], 'name')
  const existing = engine.getTable(name)
  if (!existing) {
    return rejectedReceipt(request, `Table ${name} does not exist`)
  }
  const range = readRange(input['range'], 'range')
  if (range.sheetName !== existing.sheetName) {
    return rejectedReceipt(request, `Table ${name} cannot be resized across sheets`)
  }
  const bounds = rangeBounds(range)
  const width = bounds.endCol - bounds.startCol + 1
  const columnNames = resizeColumnNames(existing.columnNames, width)
  const columns = columnNames.map((columnName, index) => ({
    ...existing.columns?.[index],
    name: columnName,
  }))
  const table: WorkbookTableSnapshot = {
    ...existing,
    startAddress: formatAddress(bounds.startRow, bounds.startCol),
    endAddress: formatAddress(bounds.endRow, bounds.endCol),
    columnNames,
    columns,
  }
  return applyTableOps(engine, request, buildSetTableOps(engine.workbook, table), {
    message: `Resize table ${table.name} to ${rangeLabel(tableRange(table))}`,
    proof: tableProof(table, 'resize'),
    changedRanges: tableSemanticRanges(table),
  })
}

function executeRenameHeader(engine: SpreadsheetEngine, request: WorkbookCommandRequest): WorkbookCommandReceipt {
  const input = inputRecord(request)
  const tableName = readNonEmptyString(input['tableName'], 'tableName')
  const table = engine.getTable(tableName)
  if (!table) {
    return rejectedReceipt(request, `Table ${tableName} does not exist`)
  }
  const nextName = readNonEmptyString(input['name'], 'name')
  const columnIndex = readColumnIndex(table, input)
  if (!table.headerRow) {
    return rejectedReceipt(request, `Table ${tableName} does not have a header row`)
  }
  const start = parseCellAddress(table.startAddress, table.sheetName)
  const address = formatAddress(start.row, start.col + columnIndex)
  return applyTableOps(
    engine,
    request,
    [
      {
        kind: 'setCellValue',
        sheetName: table.sheetName,
        address,
        value: nextName,
      },
    ],
    {
      message: `Rename ${tableName} header ${table.columnNames[columnIndex] ?? String(columnIndex + 1)} to ${nextName}`,
      proof: {
        action: 'renameHeader',
        tableName,
        columnIndex,
        columnName: nextName,
        address,
      },
      changedRanges: [{ sheetName: table.sheetName, startAddress: address, endAddress: address }, tableRange(table)],
    },
  )
}

function applyTableOps(
  engine: SpreadsheetEngine,
  request: WorkbookCommandRequest,
  ops: readonly EngineOp[],
  details: {
    readonly message: string
    readonly proof: WorkbookActionInput
    readonly changedRanges: readonly CellRangeRef[]
  },
): WorkbookCommandReceipt {
  if (ops.length === 0) {
    return {
      status: 'noop',
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      commandId: request.commandId,
      category: request.category ?? 'operation',
      previewOps: [],
      appliedOps: [],
      changedRanges: details.changedRanges,
      proof: details.proof,
      message: details.message,
    }
  }
  if (request.mode === 'preview') {
    return {
      status: 'previewed',
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      commandId: request.commandId,
      category: request.category ?? 'command',
      previewOps: ops,
      changedRanges: details.changedRanges,
      proof: details.proof,
      message: details.message,
      metadata: { mode: 'preview' },
    }
  }
  const undoOps = engine.applyOps(ops, { captureUndo: true })
  return {
    status: 'applied',
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    commandId: request.commandId,
    category: request.category ?? 'command',
    previewOps: ops,
    appliedOps: ops,
    ...(undoOps ? { undo: { id: `${request.commandId}:undo`, ops: undoOps } } : {}),
    changedRanges: details.changedRanges,
    proof: details.proof,
    message: details.message,
    metadata: { mode: request.mode ?? 'apply' },
  }
}

function projectTableRangeChrome(engine: SpreadsheetEngine, requestedRange: CellRangeRef): readonly WorkbookRangeChromeProjection[] {
  return engine
    .getTables()
    .filter((table) => rangesIntersect(tableRange(table), requestedRange))
    .flatMap((table) => tableChrome(table))
}

function tableChrome(table: WorkbookTableSnapshot): readonly WorkbookRangeChromeProjection[] {
  const range = tableRange(table)
  const bounds = rangeBounds(range)
  const headerRange = table.headerRow
    ? {
        sheetName: range.sheetName,
        startAddress: formatAddress(bounds.startRow, bounds.startCol),
        endAddress: formatAddress(bounds.startRow, bounds.endCol),
      }
    : null
  const dataStartRow = bounds.startRow + (table.headerRow ? 1 : 0)
  const dataEndRow = bounds.endRow - (table.totalsRow ? 1 : 0)
  const dataRange =
    dataStartRow <= dataEndRow
      ? {
          sheetName: range.sheetName,
          startAddress: formatAddress(dataStartRow, bounds.startCol),
          endAddress: formatAddress(dataEndRow, bounds.endCol),
        }
      : null
  const totalsRange = table.totalsRow
    ? {
        sheetName: range.sheetName,
        startAddress: formatAddress(bounds.endRow, bounds.startCol),
        endAddress: formatAddress(bounds.endRow, bounds.endCol),
      }
    : null
  return [
    chromeProjection(table, range, 'table', table.name),
    ...(headerRange ? [chromeProjection(table, headerRange, 'header', `${table.name} headers`)] : []),
    ...(dataRange ? [chromeProjection(table, dataRange, 'dataBody', `${table.name} data`)] : []),
    ...(totalsRange ? [chromeProjection(table, totalsRange, 'totals', `${table.name} totals`)] : []),
  ]
}

function chromeProjection(table: WorkbookTableSnapshot, range: CellRangeRef, role: string, label: string): WorkbookRangeChromeProjection {
  return {
    id: `tables:${table.name}:${role}`,
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    source: 'workbook-metadata',
    range,
    role,
    label,
    metadata: {
      tableName: table.name,
      sheetName: table.sheetName,
      source: 'workbook-metadata',
    },
  }
}

function tableCommandMetadata(request: WorkbookCommandRequest) {
  if (request.featureId !== WORKBOOK_TABLES_FEATURE_ID) {
    return undefined
  }
  return {
    label: request.commandId,
    metadata: {
      featureId: WORKBOOK_TABLES_FEATURE_ID,
      commandId: request.commandId,
      source: 'tables-feature',
    },
  }
}

function inputRecord(request: WorkbookCommandRequest): Record<string, unknown> {
  if (!isRecord(request.input)) {
    throw new Error(`Workbook command ${request.commandId} requires object input`)
  }
  return request.input
}

function readRange(value: unknown, label: string): CellRangeRef {
  if (!isCellRangeLike(value)) {
    throw new Error(`Table command ${label} must be a workbook range`)
  }
  return {
    sheetName: value.sheetName,
    startAddress: value.startAddress,
    endAddress: value.endAddress,
  }
}

function readTable(value: unknown): WorkbookTableSnapshot {
  if (!isRecord(value)) {
    throw new Error('Table command table must be an object')
  }
  return {
    name: readNonEmptyString(value['name'], 'table.name'),
    sheetName: readNonEmptyString(value['sheetName'], 'table.sheetName'),
    startAddress: readNonEmptyString(value['startAddress'], 'table.startAddress'),
    endAddress: readNonEmptyString(value['endAddress'], 'table.endAddress'),
    columnNames: readStringList(value['columnNames'], 'table.columnNames'),
    headerRow: readBoolean(value['headerRow'], 'table.headerRow'),
    totalsRow: readBoolean(value['totalsRow'], 'table.totalsRow'),
    ...(value['columns'] !== undefined ? { columns: readTableColumns(value['columns']) } : {}),
    ...(value['style'] !== undefined ? { style: readTableStyle(value['style']) } : {}),
    ...(value['autoFilter'] !== undefined ? { autoFilter: readAutoFilter(value['autoFilter']) } : {}),
    ...(typeof value['sortState'] === 'string' ? { sortState: value['sortState'] } : {}),
  }
}

function readTableColumns(value: unknown): NonNullable<WorkbookTableSnapshot['columns']> {
  if (!Array.isArray(value)) {
    throw new Error('Table command table.columns must be a list')
  }
  return value.map((column) => {
    if (!isRecord(column)) {
      throw new Error('Table column must be an object')
    }
    return {
      name: readNonEmptyString(column['name'], 'table.columns.name'),
      ...(typeof column['calculatedColumnFormula'] === 'string' ? { calculatedColumnFormula: column['calculatedColumnFormula'] } : {}),
      ...(typeof column['totalsRowLabel'] === 'string' ? { totalsRowLabel: column['totalsRowLabel'] } : {}),
      ...(typeof column['totalsRowFunction'] === 'string' ? { totalsRowFunction: column['totalsRowFunction'] } : {}),
      ...(typeof column['totalsRowFormula'] === 'string' ? { totalsRowFormula: column['totalsRowFormula'] } : {}),
    }
  })
}

function readTableStyle(value: unknown): NonNullable<WorkbookTableSnapshot['style']> {
  if (!isRecord(value)) {
    throw new Error('Table command table.style must be an object')
  }
  return {
    ...(typeof value['name'] === 'string' ? { name: value['name'] } : {}),
    ...(typeof value['showFirstColumn'] === 'boolean' ? { showFirstColumn: value['showFirstColumn'] } : {}),
    ...(typeof value['showLastColumn'] === 'boolean' ? { showLastColumn: value['showLastColumn'] } : {}),
    ...(typeof value['showRowStripes'] === 'boolean' ? { showRowStripes: value['showRowStripes'] } : {}),
    ...(typeof value['showColumnStripes'] === 'boolean' ? { showColumnStripes: value['showColumnStripes'] } : {}),
  }
}

function readAutoFilter(value: unknown): NonNullable<WorkbookTableSnapshot['autoFilter']> {
  const range = readRange(value, 'table.autoFilter')
  if (!isRecord(value)) {
    throw new Error('Table command table.autoFilter must be an object')
  }
  return {
    ...range,
    ...(Array.isArray(value['criteria']) ? { criteria: structuredClone(value['criteria']) } : {}),
  }
}

function readColumnIndex(table: WorkbookTableSnapshot, input: Record<string, unknown>): number {
  if (typeof input['columnIndex'] === 'number' && Number.isSafeInteger(input['columnIndex']) && input['columnIndex'] >= 0) {
    if (input['columnIndex'] >= table.columnNames.length) {
      throw new Error(`Table ${table.name} column index is out of bounds`)
    }
    return input['columnIndex']
  }
  const columnName = readNonEmptyString(input['columnName'], 'columnName')
  const columnIndex = table.columnNames.findIndex((name) => name.trim().toUpperCase() === columnName.trim().toUpperCase())
  if (columnIndex < 0) {
    throw new Error(`Table ${table.name} column ${columnName} does not exist`)
  }
  return columnIndex
}

function readNonEmptyString(value: unknown, label: string): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`Table command ${label} must be a non-empty string`)
  }
  return value.trim()
}

function readStringList(value: unknown, label: string): string[] {
  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    throw new Error(`Table command ${label} must be a string list`)
  }
  return value.map((entry) => entry.trim())
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Table command ${label} must be a boolean`)
  }
  return value
}

function isCellRangeLike(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function rangeBounds(range: CellRangeRef): {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
} {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    sheetName: range.sheetName,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function detectHeaderRow(engine: SpreadsheetEngine, bounds: ReturnType<typeof rangeBounds>): boolean {
  for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
    if (cellValueText(engine.getCellValue(bounds.sheetName, formatAddress(bounds.startRow, col))).trim() !== '') {
      return true
    }
  }
  return false
}

function tableColumnNamesForRange(engine: SpreadsheetEngine, bounds: ReturnType<typeof rangeBounds>, hasHeaders: boolean): string[] {
  const names: string[] = []
  for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
    const requested = hasHeaders ? cellValueText(engine.getCellValue(bounds.sheetName, formatAddress(bounds.startRow, col))) : ''
    names.push(excelCompatibleTableColumnName(requested, names, names.length))
  }
  return names
}

function resizeColumnNames(previous: readonly string[], width: number): string[] {
  const names: string[] = []
  for (let index = 0; index < width; index += 1) {
    names.push(excelCompatibleTableColumnName(previous[index] ?? '', names, index))
  }
  return names
}

function cellValueText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.String:
      return value.value
    case ValueTag.Number:
      return String(value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.Error:
      return '#ERROR'
    case ValueTag.Empty:
      return ''
  }
}

function nextTableName(engine: SpreadsheetEngine): string {
  const existing = new Set(engine.getTables().map((table) => table.name.trim().toUpperCase()))
  let suffix = 1
  while (existing.has(`TABLE${String(suffix)}`)) {
    suffix += 1
  }
  return `Table${String(suffix)}`
}

function tableRange(table: WorkbookTableSnapshot): CellRangeRef {
  return {
    sheetName: table.sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
  }
}

function tableSemanticRanges(table: WorkbookTableSnapshot): readonly CellRangeRef[] {
  return [tableRange(table)]
}

function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  if (left.sheetName !== right.sheetName) {
    return false
  }
  const leftBounds = rangeBounds(left)
  const rightBounds = rangeBounds(right)
  return !(
    leftBounds.endRow < rightBounds.startRow ||
    rightBounds.endRow < leftBounds.startRow ||
    leftBounds.endCol < rightBounds.startCol ||
    rightBounds.endCol < leftBounds.startCol
  )
}

function tableProof(table: WorkbookTableSnapshot, action: string): WorkbookActionInput {
  return {
    action,
    tableName: table.name,
    sheetName: table.sheetName,
    range: rangeActionInput(tableRange(table)),
    columnNames: [...table.columnNames],
    source: 'workbook-command-service',
  }
}

function rangeActionInput(range: CellRangeRef): WorkbookActionInput {
  return {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }
}

function rangeLabel(range: CellRangeRef): string {
  return `${range.sheetName}!${range.startAddress}:${range.endAddress}`
}

function rejectedReceipt(request: WorkbookCommandRequest, message: string): WorkbookCommandReceipt {
  return {
    status: 'rejected',
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    commandId: request.commandId,
    category: request.category ?? 'command',
    message,
    errors: [message],
  }
}

function noopReceipt(request: WorkbookCommandRequest, message: string): WorkbookCommandReceipt {
  return {
    status: 'noop',
    featureId: WORKBOOK_TABLES_FEATURE_ID,
    commandId: request.commandId,
    category: request.category ?? 'operation',
    previewOps: [],
    appliedOps: [],
    message,
  }
}
