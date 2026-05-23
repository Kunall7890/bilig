import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from './workbook-range-records.js'
import {
  cloneCommentThreadRecord,
  cloneHyperlinkRecord,
  cloneNoteRecord,
  clonePivotRecord,
  cloneSpillRecord,
  commentThreadKey,
  hyperlinkKey,
  noteKey,
  spillKey,
} from './workbook-metadata-records.js'
import {
  pivotKey,
  type WorkbookCommentThreadRecord,
  type WorkbookHyperlinkRecord,
  type WorkbookMetadataRecord,
  type WorkbookNoteRecord,
  type WorkbookPivotRecord,
  type WorkbookSpillRecord,
} from './workbook-metadata-types.js'
import type { WorkbookMetadataService } from './workbook-metadata-service-contract.js'
import { metadataEffect } from './workbook-metadata-service-helpers.js'

type WorkbookMetadataCellRecordService = Pick<
  WorkbookMetadataService,
  | 'setCommentThread'
  | 'getCommentThread'
  | 'deleteCommentThread'
  | 'listCommentThreads'
  | 'setNote'
  | 'getNote'
  | 'deleteNote'
  | 'listNotes'
  | 'setHyperlink'
  | 'getHyperlink'
  | 'deleteHyperlink'
  | 'listHyperlinks'
  | 'setSpill'
  | 'getSpill'
  | 'deleteSpill'
  | 'listSpills'
  | 'setPivot'
  | 'getPivot'
  | 'getPivotByKey'
  | 'deletePivot'
  | 'hasPivots'
  | 'listPivots'
>

export function createWorkbookMetadataCellRecordService(metadata: WorkbookMetadataRecord): WorkbookMetadataCellRecordService {
  return {
    setCommentThread(record) {
      return metadataEffect('Failed to set comment thread metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookCommentThreadRecord = cloneCommentThreadRecord({
          ...record,
          threadId: record.threadId.trim(),
          address: normalizedAddress,
          comments: record.comments.map((comment) => ({
            id: comment.id.trim(),
            body: comment.body.trim(),
            ...(comment.authorUserId !== undefined ? { authorUserId: comment.authorUserId } : {}),
            ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
            ...(comment.createdAtUnixMs !== undefined ? { createdAtUnixMs: comment.createdAtUnixMs } : {}),
          })),
        })
        metadata.commentThreads.set(commentThreadKey(record.sheetName, normalizedAddress), stored)
        return cloneCommentThreadRecord(stored)
      })
    },
    getCommentThread(sheetName, address) {
      return metadataEffect('Failed to get comment thread metadata', () => {
        const record = metadata.commentThreads.get(commentThreadKey(sheetName, address))
        return record ? cloneCommentThreadRecord(record) : undefined
      })
    },
    deleteCommentThread(sheetName, address) {
      return metadataEffect('Failed to delete comment thread metadata', () =>
        metadata.commentThreads.delete(commentThreadKey(sheetName, address)),
      )
    },
    listCommentThreads(sheetName) {
      return metadataEffect('Failed to list comment thread metadata', () =>
        [...metadata.commentThreads.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) =>
            commentThreadKey(left.sheetName, left.address).localeCompare(commentThreadKey(right.sheetName, right.address)),
          )
          .map(cloneCommentThreadRecord),
      )
    },
    setNote(record) {
      return metadataEffect('Failed to set note metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookNoteRecord = cloneNoteRecord({
          sheetName: record.sheetName,
          address: normalizedAddress,
          text: record.text.trim(),
        })
        metadata.notes.set(noteKey(record.sheetName, normalizedAddress), stored)
        return cloneNoteRecord(stored)
      })
    },
    getNote(sheetName, address) {
      return metadataEffect('Failed to get note metadata', () => {
        const record = metadata.notes.get(noteKey(sheetName, address))
        return record ? cloneNoteRecord(record) : undefined
      })
    },
    deleteNote(sheetName, address) {
      return metadataEffect('Failed to delete note metadata', () => metadata.notes.delete(noteKey(sheetName, address)))
    },
    listNotes(sheetName) {
      return metadataEffect('Failed to list note metadata', () =>
        [...metadata.notes.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => noteKey(left.sheetName, left.address).localeCompare(noteKey(right.sheetName, right.address)))
          .map(cloneNoteRecord),
      )
    },
    setHyperlink(record) {
      return metadataEffect('Failed to set hyperlink metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookHyperlinkRecord = cloneHyperlinkRecord({
          sheetName: record.sheetName,
          address: normalizedAddress,
          target: record.target,
          ...(record.tooltip !== undefined ? { tooltip: record.tooltip } : {}),
          ...(record.display !== undefined ? { display: record.display } : {}),
        })
        metadata.hyperlinks.set(hyperlinkKey(record.sheetName, normalizedAddress), stored)
        return cloneHyperlinkRecord(stored)
      })
    },
    getHyperlink(sheetName, address) {
      return metadataEffect('Failed to get hyperlink metadata', () => {
        const record = metadata.hyperlinks.get(hyperlinkKey(sheetName, address))
        return record ? cloneHyperlinkRecord(record) : undefined
      })
    },
    deleteHyperlink(sheetName, address) {
      return metadataEffect('Failed to delete hyperlink metadata', () => metadata.hyperlinks.delete(hyperlinkKey(sheetName, address)))
    },
    listHyperlinks(sheetName) {
      return metadataEffect('Failed to list hyperlink metadata', () =>
        [...metadata.hyperlinks.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => hyperlinkKey(left.sheetName, left.address).localeCompare(hyperlinkKey(right.sheetName, right.address)))
          .map(cloneHyperlinkRecord),
      )
    },
    setSpill(sheetName, address, rows, cols) {
      return metadataEffect('Failed to set spill metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(sheetName, address)
        const record: WorkbookSpillRecord = { sheetName, address: normalizedAddress, rows, cols }
        metadata.spills.set(spillKey(sheetName, normalizedAddress), record)
        return { ...record }
      })
    },
    getSpill(sheetName, address) {
      return metadataEffect('Failed to get spill metadata', () => {
        const record = metadata.spills.get(spillKey(sheetName, address))
        return record ? { ...record } : undefined
      })
    },
    deleteSpill(sheetName, address) {
      return metadataEffect('Failed to delete spill metadata', () => metadata.spills.delete(spillKey(sheetName, address)))
    },
    listSpills() {
      return metadataEffect('Failed to list spill metadata', () =>
        [...metadata.spills.values()]
          .toSorted((left, right) => `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`))
          .map(cloneSpillRecord),
      )
    },
    setPivot(record) {
      return metadataEffect('Failed to set pivot metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookPivotRecord = {
          ...record,
          name: record.name.trim(),
          address: normalizedAddress,
          groupBy: [...record.groupBy],
          values: record.values.map((value) => ({ ...value })),
          ...(record.source ? { source: canonicalWorkbookRangeRef(record.source) } : {}),
        }
        metadata.pivots.set(pivotKey(record.sheetName, normalizedAddress), stored)
        return clonePivotRecord(stored)
      })
    },
    getPivot(sheetName, address) {
      return metadataEffect('Failed to get pivot metadata', () => {
        const record = metadata.pivots.get(pivotKey(sheetName, address))
        return record ? clonePivotRecord(record) : undefined
      })
    },
    getPivotByKey(key) {
      return metadataEffect('Failed to get pivot metadata by key', () => {
        const record = metadata.pivots.get(key)
        return record ? clonePivotRecord(record) : undefined
      })
    },
    deletePivot(sheetName, address) {
      return metadataEffect('Failed to delete pivot metadata', () => metadata.pivots.delete(pivotKey(sheetName, address)))
    },
    hasPivots() {
      return metadataEffect('Failed to read pivot metadata state', () => metadata.pivots.size > 0)
    },
    listPivots() {
      return metadataEffect('Failed to list pivot metadata', () =>
        [...metadata.pivots.values()]
          .toSorted((left, right) => `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`))
          .map(clonePivotRecord),
      )
    },
  }
}
