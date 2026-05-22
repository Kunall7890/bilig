import type {
  CellRangeRef,
  LiteralInput,
  SheetFormatRangeSnapshot,
  SheetStyleRangeSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookNoteSnapshot,
  WorkbookSnapshot,
  WorkbookSortSnapshot,
} from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'
import { WORKBOOK_DEFAULT_FORMAT_ID, WORKBOOK_DEFAULT_STYLE_ID } from '../workbook-default-style-format.js'

export interface SpreadsheetEngineSortRangeOptions {
  readonly header?: boolean
}

interface SortBounds {
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

interface SortCell {
  readonly formula?: string
  readonly format?: string
  readonly value?: LiteralInput
}

interface SortRow {
  readonly row: number
  readonly originalIndex: number
}

interface SortCellRewrite {
  readonly col: number
  readonly sourceRow: number
  readonly targetRow: number
  readonly sourceAddress: string
  readonly targetAddress: string
}

type RangeStyleFormatSnapshot = SheetStyleRangeSnapshot | SheetFormatRangeSnapshot

export function buildSortRangeOps(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  range: CellRangeRef,
  keys: readonly WorkbookSortSnapshot['keys'][number][],
  options: SpreadsheetEngineSortRangeOptions = {},
): EngineOp[] {
  if (keys.length === 0) {
    throw new Error('sortRange requires at least one sort key')
  }

  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  if (!sheet) {
    throw new Error(`Unknown sheet: ${sheetName}`)
  }

  const bounds = sortBounds(sheetName, range)
  const bodyStartRow = options.header === true ? bounds.startRow + 1 : bounds.startRow
  if (bodyStartRow > bounds.endRow) {
    return []
  }

  const cellsByAddress = new Map<string, SortCell>()
  for (const cell of sheet.cells) {
    cellsByAddress.set(cell.address, {
      ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
      ...(cell.value !== undefined ? { value: cell.value } : {}),
      ...(cell.format !== undefined ? { format: cell.format } : {}),
    })
  }

  const keyColumns = keys.map((key) => parseCellAddress(key.keyAddress, sheetName).col)
  const rows = Array.from({ length: bounds.endRow - bodyStartRow + 1 }, (_, index): SortRow => {
    return { row: bodyStartRow + index, originalIndex: index }
  })
  const sortedRows = rows.toSorted((left, right) => compareSortRows(cellsByAddress, keyColumns, keys, left, right))

  const ops: EngineOp[] = []
  const rewrites = buildSortCellRewrites(sortedRows, bounds, bodyStartRow)
  for (const rewrite of rewrites) {
    const source = cellsByAddress.get(rewrite.sourceAddress)
    const existing = cellsByAddress.get(rewrite.targetAddress)
    appendCellReplacementOps(ops, sheetName, rewrite.targetAddress, existing, source)
  }
  appendRowBundleMetadataSortOps(ops, sheet.metadata, rewrites, sheetName)

  const metadataRange: CellRangeRef = {
    sheetName,
    startAddress: formatAddress(bodyStartRow, bounds.startCol),
    endAddress: formatAddress(bounds.endRow, bounds.endCol),
  }
  ops.push({
    kind: 'setSort',
    sheetName,
    range: metadataRange,
    keys: keys.map((key) => ({ keyAddress: key.keyAddress, direction: key.direction })),
  })
  return ops
}

function buildSortCellRewrites(sortedRows: readonly SortRow[], bounds: SortBounds, bodyStartRow: number): SortCellRewrite[] {
  return sortedRows.flatMap((sortRow, targetOffset) => {
    const targetRow = bodyStartRow + targetOffset
    return Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, index): SortCellRewrite => {
      const col = bounds.startCol + index
      return {
        col,
        sourceRow: sortRow.row,
        targetRow,
        sourceAddress: formatAddress(sortRow.row, col),
        targetAddress: formatAddress(targetRow, col),
      }
    })
  })
}

function appendRowBundleMetadataSortOps(
  ops: EngineOp[],
  metadata: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']> | undefined,
  rewrites: readonly SortCellRewrite[],
  sheetName: string,
): void {
  if (!metadata) {
    return
  }
  appendStyleRangeSortOps(ops, metadata.styleRanges, rewrites, sheetName)
  appendFormatRangeSortOps(ops, metadata.formatRanges, rewrites, sheetName)
  appendCommentThreadSortOps(ops, metadata.commentThreads, rewrites, sheetName)
  appendNoteSortOps(ops, metadata.notes, rewrites, sheetName)
}

function appendCommentThreadSortOps(
  ops: EngineOp[],
  commentThreads: readonly WorkbookCommentThreadSnapshot[] | undefined,
  rewrites: readonly SortCellRewrite[],
  sheetName: string,
): void {
  const threadsByAddress = new Map((commentThreads ?? []).map((thread) => [thread.address, thread]))
  for (const rewrite of rewrites) {
    const source = threadsByAddress.get(rewrite.sourceAddress)
    const existing = threadsByAddress.get(rewrite.targetAddress)
    if (source) {
      ops.push({
        kind: 'upsertCommentThread',
        thread: {
          ...structuredClone(source),
          sheetName,
          address: rewrite.targetAddress,
        },
      })
    } else if (existing) {
      ops.push({ kind: 'deleteCommentThread', sheetName, address: rewrite.targetAddress })
    }
  }
}

function appendNoteSortOps(
  ops: EngineOp[],
  notes: readonly WorkbookNoteSnapshot[] | undefined,
  rewrites: readonly SortCellRewrite[],
  sheetName: string,
): void {
  const notesByAddress = new Map((notes ?? []).map((note) => [note.address, note]))
  for (const rewrite of rewrites) {
    const source = notesByAddress.get(rewrite.sourceAddress)
    const existing = notesByAddress.get(rewrite.targetAddress)
    if (source) {
      ops.push({
        kind: 'upsertNote',
        note: {
          ...structuredClone(source),
          sheetName,
          address: rewrite.targetAddress,
        },
      })
    } else if (existing) {
      ops.push({ kind: 'deleteNote', sheetName, address: rewrite.targetAddress })
    }
  }
}

function appendStyleRangeSortOps(
  ops: EngineOp[],
  styleRanges: readonly SheetStyleRangeSnapshot[] | undefined,
  rewrites: readonly SortCellRewrite[],
  sheetName: string,
): void {
  if ((styleRanges?.length ?? 0) === 0) {
    return
  }
  appendRangeIdSortOps(
    rewrites,
    WORKBOOK_DEFAULT_STYLE_ID,
    (row, col) => rangeStyleIdAt(styleRanges, row, col),
    (range, styleId) => ops.push({ kind: 'setStyleRange', range: { ...range, sheetName }, styleId }),
  )
}

function appendFormatRangeSortOps(
  ops: EngineOp[],
  formatRanges: readonly SheetFormatRangeSnapshot[] | undefined,
  rewrites: readonly SortCellRewrite[],
  sheetName: string,
): void {
  if ((formatRanges?.length ?? 0) === 0) {
    return
  }
  appendRangeIdSortOps(
    rewrites,
    WORKBOOK_DEFAULT_FORMAT_ID,
    (row, col) => rangeFormatIdAt(formatRanges, row, col),
    (range, formatId) => ops.push({ kind: 'setFormatRange', range: { ...range, sheetName }, formatId }),
  )
}

function appendRangeIdSortOps(
  rewrites: readonly SortCellRewrite[],
  defaultId: string,
  idAt: (row: number, col: number) => string | undefined,
  appendOp: (range: Omit<CellRangeRef, 'sheetName'>, id: string) => void,
): void {
  let pending: { targetRow: number; startCol: number; endCol: number; id: string } | undefined
  const flush = (): void => {
    if (!pending) {
      return
    }
    appendOp(
      {
        startAddress: formatAddress(pending.targetRow, pending.startCol),
        endAddress: formatAddress(pending.targetRow, pending.endCol),
      },
      pending.id,
    )
    pending = undefined
  }

  for (const rewrite of rewrites) {
    const sourceId = idAt(rewrite.sourceRow, rewrite.col) ?? defaultId
    const targetId = idAt(rewrite.targetRow, rewrite.col) ?? defaultId
    if (sourceId === targetId) {
      flush()
      continue
    }
    if (pending && pending.targetRow === rewrite.targetRow && pending.endCol + 1 === rewrite.col && pending.id === sourceId) {
      pending = { ...pending, endCol: rewrite.col }
      continue
    }
    flush()
    pending = {
      targetRow: rewrite.targetRow,
      startCol: rewrite.col,
      endCol: rewrite.col,
      id: sourceId,
    }
  }
  flush()
}

function rangeStyleIdAt(styleRanges: readonly SheetStyleRangeSnapshot[] | undefined, row: number, col: number): string | undefined {
  return rangeRecordAt(styleRanges, row, col, (record) => record.styleId)
}

function rangeFormatIdAt(formatRanges: readonly SheetFormatRangeSnapshot[] | undefined, row: number, col: number): string | undefined {
  return rangeRecordAt(formatRanges, row, col, (record) => record.formatId)
}

function rangeRecordAt<T extends RangeStyleFormatSnapshot>(
  records: readonly T[] | undefined,
  row: number,
  col: number,
  readId: (record: T) => string,
): string | undefined {
  if (!records) {
    return undefined
  }
  for (let index = records.length - 1; index >= 0; index -= 1) {
    const record = records[index]!
    if (rangeContainsCell(record.range, row, col)) {
      return readId(record)
    }
  }
  return undefined
}

function rangeContainsCell(range: CellRangeRef, row: number, col: number): boolean {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return (
    row >= Math.min(start.row, end.row) &&
    row <= Math.max(start.row, end.row) &&
    col >= Math.min(start.col, end.col) &&
    col <= Math.max(start.col, end.col)
  )
}

function sortBounds(sheetName: string, range: CellRangeRef): SortBounds {
  const start = parseCellAddress(range.startAddress, sheetName)
  const end = parseCellAddress(range.endAddress, sheetName)
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function compareSortRows(
  cellsByAddress: ReadonlyMap<string, SortCell>,
  keyColumns: readonly number[],
  keys: readonly WorkbookSortSnapshot['keys'][number][],
  left: SortRow,
  right: SortRow,
): number {
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index]!
    const col = keyColumns[index]!
    const leftValue = cellsByAddress.get(formatAddress(left.row, col))?.value
    const rightValue = cellsByAddress.get(formatAddress(right.row, col))?.value
    const compared = compareSortValues(leftValue, rightValue, key.direction)
    if (compared !== 0) {
      return compared
    }
  }
  return left.originalIndex - right.originalIndex
}

function compareSortValues(left: LiteralInput | undefined, right: LiteralInput | undefined, direction: 'asc' | 'desc'): number {
  const leftBlank = left === undefined || left === null
  const rightBlank = right === undefined || right === null
  if (leftBlank || rightBlank) {
    if (leftBlank === rightBlank) {
      return 0
    }
    return leftBlank ? 1 : -1
  }

  const leftRank = sortValueRank(left)
  const rightRank = sortValueRank(right)
  const compared =
    leftRank !== rightRank
      ? leftRank - rightRank
      : typeof left === 'number' && typeof right === 'number'
        ? left - right
        : typeof left === 'boolean' && typeof right === 'boolean'
          ? Number(left) - Number(right)
          : String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: 'base' })
  return direction === 'desc' ? -compared : compared
}

function sortValueRank(value: LiteralInput | undefined): number {
  if (value === undefined || value === null) {
    return 3
  }
  if (typeof value === 'number') {
    return 0
  }
  if (typeof value === 'string') {
    return 1
  }
  if (typeof value === 'boolean') {
    return 2
  }
  return 3
}

function appendCellReplacementOps(
  ops: EngineOp[],
  sheetName: string,
  address: string,
  existing: SortCell | undefined,
  source: SortCell | undefined,
): void {
  if (!source) {
    if (existing) {
      ops.push({ kind: 'clearCell', sheetName, address })
      if (existing.format !== undefined) {
        ops.push({ kind: 'setCellFormat', sheetName, address, format: null })
      }
    }
    return
  }

  if (source.formula !== undefined) {
    if (existing?.formula !== source.formula) {
      ops.push({ kind: 'setCellFormula', sheetName, address, formula: source.formula })
    }
  } else if (source.value !== undefined) {
    if (existing?.formula !== undefined || existing?.value !== source.value) {
      ops.push({ kind: 'setCellValue', sheetName, address, value: source.value })
    }
  } else if (existing?.formula !== undefined || existing?.value !== undefined) {
    ops.push({ kind: 'clearCell', sheetName, address })
  }

  if (existing?.format !== source.format) {
    ops.push({ kind: 'setCellFormat', sheetName, address, format: source.format ?? null })
  }
}
