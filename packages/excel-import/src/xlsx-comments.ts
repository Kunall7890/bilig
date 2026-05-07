import * as XLSX from 'xlsx'

import type { WorkbookCommentThreadSnapshot } from '@bilig/protocol'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorksheetCellObject(value: unknown): value is XLSX.CellObject {
  return isRecord(value) && typeof value['t'] === 'string'
}

function normalizeCommentAddress(value: string): string | null {
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(value))
  } catch {
    return null
  }
}

function denseWorksheetRows(sheet: XLSX.WorkSheet): unknown[] | null {
  const denseRows = (sheet as Record<string, unknown>)['!data']
  return Array.isArray(denseRows) ? denseRows : null
}

function readCellComments(value: unknown): unknown[] | null {
  return isRecord(value) && Array.isArray(value['c']) && value['c'].length > 0 ? value['c'] : null
}

function appendCommentThread(input: {
  commentThreads: WorkbookCommentThreadSnapshot[]
  sheetName: string
  address: string
  commentsValue: readonly unknown[]
  ignoredCount: number
}): number {
  const threadId = `xlsx-comment:${input.sheetName}:${input.address}`
  let ignoredCount = input.ignoredCount
  const comments = input.commentsValue.flatMap((comment, index) => {
    if (!isRecord(comment) || typeof comment['t'] !== 'string') {
      ignoredCount += 1
      return []
    }
    const authorDisplayName = typeof comment['a'] === 'string' && comment['a'].trim().length > 0 ? comment['a'].trim() : undefined
    return [
      {
        id: `${threadId}:${index + 1}`,
        body: comment['t'],
        ...(authorDisplayName !== undefined ? { authorDisplayName } : {}),
      },
    ]
  })
  if (comments.length > 0) {
    input.commentThreads.push({
      threadId,
      sheetName: input.sheetName,
      address: input.address,
      comments,
    })
  }
  return ignoredCount
}

export function readImportedSheetComments(
  sheetName: string,
  sheet: XLSX.WorkSheet,
): {
  commentThreads: WorkbookCommentThreadSnapshot[] | undefined
  ignoredCount: number
} {
  const commentThreads: WorkbookCommentThreadSnapshot[] = []
  let ignoredCount = 0

  const denseRows = denseWorksheetRows(sheet)
  if (denseRows) {
    denseRows.forEach((row, rowIndex) => {
      if (!Array.isArray(row)) {
        return
      }
      row.forEach((cell, columnIndex) => {
        const commentsValue = readCellComments(cell)
        if (!commentsValue) {
          return
        }
        ignoredCount = appendCommentThread({
          commentThreads,
          sheetName,
          address: XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex }),
          commentsValue,
          ignoredCount,
        })
      })
    })
  }

  for (const [key, value] of Object.entries(sheet)) {
    if (key.startsWith('!') || !isRecord(value)) {
      continue
    }
    const address = normalizeCommentAddress(key)
    const commentsValue = readCellComments(value)
    if (!address || !commentsValue) {
      continue
    }

    ignoredCount = appendCommentThread({
      commentThreads,
      sheetName,
      address,
      commentsValue,
      ignoredCount,
    })
  }

  commentThreads.sort((left, right) =>
    `${left.sheetName}:${left.address}:${left.threadId}`.localeCompare(`${right.sheetName}:${right.address}:${right.threadId}`),
  )
  return {
    commentThreads: commentThreads.length > 0 ? commentThreads : undefined,
    ignoredCount,
  }
}

export function addExportCommentsToWorksheet(
  worksheet: XLSX.WorkSheet,
  commentThreads: readonly WorkbookCommentThreadSnapshot[] | undefined,
): void {
  if (!commentThreads || commentThreads.length === 0) {
    return
  }

  for (const thread of commentThreads) {
    const address = normalizeCommentAddress(thread.address)
    if (!address || thread.comments.length === 0) {
      continue
    }
    const existingCell = worksheet[address]
    const cell = isWorksheetCellObject(existingCell) ? existingCell : ({ t: 'z' } satisfies XLSX.CellObject)
    const comments: XLSX.Comment[] = []
    for (const comment of thread.comments) {
      const xlsxComment: XLSX.Comment = { t: comment.body }
      if (comment.authorDisplayName !== undefined) {
        xlsxComment.a = comment.authorDisplayName
      }
      comments.push(xlsxComment)
    }
    cell.c = comments
    worksheet[address] = cell
  }
}
