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

export function readImportedSheetComments(
  sheetName: string,
  sheet: XLSX.WorkSheet,
): {
  commentThreads: WorkbookCommentThreadSnapshot[] | undefined
  ignoredCount: number
} {
  const commentThreads: WorkbookCommentThreadSnapshot[] = []
  let ignoredCount = 0

  for (const [key, value] of Object.entries(sheet)) {
    if (key.startsWith('!') || !isRecord(value)) {
      continue
    }
    const address = normalizeCommentAddress(key)
    const commentsValue = value['c']
    if (!address || !Array.isArray(commentsValue) || commentsValue.length === 0) {
      continue
    }

    const threadId = `xlsx-comment:${sheetName}:${address}`
    const comments = commentsValue.flatMap((comment, index) => {
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
    if (comments.length === 0) {
      continue
    }
    commentThreads.push({
      threadId,
      sheetName,
      address,
      comments,
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
