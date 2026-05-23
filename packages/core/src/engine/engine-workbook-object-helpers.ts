import type {
  WorkbookChartSnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookHyperlinkSnapshot,
  WorkbookImageSnapshot,
  WorkbookNoteSnapshot,
  WorkbookShapeSnapshot,
} from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkbookTableRecord } from '../workbook-store.js'

export function normalizeEngineCommentThread(thread: WorkbookCommentThreadSnapshot): WorkbookCommentThreadSnapshot {
  const parsed = parseCellAddress(thread.address, thread.sheetName)
  return {
    threadId: thread.threadId.trim(),
    sheetName: thread.sheetName,
    address: formatAddress(parsed.row, parsed.col),
    comments: thread.comments.map((comment) => ({
      id: comment.id.trim(),
      body: comment.body.trim(),
      ...(comment.authorUserId !== undefined ? { authorUserId: comment.authorUserId } : {}),
      ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
      ...(comment.createdAtUnixMs !== undefined ? { createdAtUnixMs: comment.createdAtUnixMs } : {}),
    })),
    ...(thread.resolved !== undefined ? { resolved: thread.resolved } : {}),
    ...(thread.resolvedByUserId !== undefined ? { resolvedByUserId: thread.resolvedByUserId } : {}),
    ...(thread.resolvedAtUnixMs !== undefined ? { resolvedAtUnixMs: thread.resolvedAtUnixMs } : {}),
  }
}

export function normalizeEngineNote(note: WorkbookNoteSnapshot): WorkbookNoteSnapshot {
  const parsed = parseCellAddress(note.address, note.sheetName)
  return {
    sheetName: note.sheetName,
    address: formatAddress(parsed.row, parsed.col),
    text: note.text.trim(),
  }
}

export function normalizeEngineHyperlink(hyperlink: WorkbookHyperlinkSnapshot): WorkbookHyperlinkSnapshot {
  const parsed = parseCellAddress(hyperlink.address, hyperlink.sheetName)
  return {
    sheetName: hyperlink.sheetName,
    address: formatAddress(parsed.row, parsed.col),
    target: hyperlink.target,
    ...(hyperlink.tooltip !== undefined ? { tooltip: hyperlink.tooltip } : {}),
    ...(hyperlink.display !== undefined ? { display: hyperlink.display } : {}),
  }
}

export function workbookObjectRecordEqual<T>(left: T | undefined, right: T): boolean {
  return left !== undefined && JSON.stringify(left) === JSON.stringify(right)
}

export function workbookTablesEqual(left: WorkbookTableRecord | undefined, right: WorkbookTableRecord): boolean {
  return workbookObjectRecordEqual(left, right)
}

export function cloneEngineTableRecord(table: WorkbookTableRecord): WorkbookTableRecord {
  return structuredClone(table)
}

export function workbookChartsEqual(left: WorkbookChartSnapshot | undefined, right: WorkbookChartSnapshot): boolean {
  return (
    left !== undefined &&
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.chartType === right.chartType &&
    left.source.sheetName === right.source.sheetName &&
    left.source.startAddress === right.source.startAddress &&
    left.source.endAddress === right.source.endAddress &&
    left.rows === right.rows &&
    left.cols === right.cols &&
    JSON.stringify(left.anchor) === JSON.stringify(right.anchor) &&
    left.seriesOrientation === right.seriesOrientation &&
    left.firstRowAsHeaders === right.firstRowAsHeaders &&
    left.firstColumnAsLabels === right.firstColumnAsLabels &&
    left.title === right.title &&
    left.legendPosition === right.legendPosition
  )
}

export function workbookImagesEqual(left: WorkbookImageSnapshot | undefined, right: WorkbookImageSnapshot): boolean {
  return (
    left !== undefined &&
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.sourceUrl === right.sourceUrl &&
    left.rows === right.rows &&
    left.cols === right.cols &&
    left.altText === right.altText
  )
}

export function workbookShapesEqual(left: WorkbookShapeSnapshot | undefined, right: WorkbookShapeSnapshot): boolean {
  return (
    left !== undefined &&
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.shapeType === right.shapeType &&
    left.rows === right.rows &&
    left.cols === right.cols &&
    left.text === right.text &&
    left.fillColor === right.fillColor &&
    left.strokeColor === right.strokeColor
  )
}
