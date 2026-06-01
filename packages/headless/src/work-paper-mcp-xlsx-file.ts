import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { basename, dirname, resolve } from 'node:path'

import { exportWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import { WorkPaper } from './work-paper.js'
import { createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import type { WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import { importXlsx } from './xlsx.js'

interface FileBackedWorkPaperMcpFromXlsxOptions {
  readonly fromXlsxPath: string
  readonly overwriteWorkPaper?: boolean
  readonly workpaperPath: string
  readonly writable?: boolean
}

function createFileBackedWorkPaperMcpToolServerFromXlsxFile(input: FileBackedWorkPaperMcpFromXlsxOptions): WorkPaperMcpToolServer {
  const workpaperPath = resolve(input.workpaperPath)
  if (existsSync(workpaperPath) && !input.overwriteWorkPaper) {
    throw new Error(`WorkPaper JSON already exists at ${workpaperPath}; pass --overwrite-workpaper to replace it`)
  }

  const xlsxPath = resolve(input.fromXlsxPath)
  const imported = importXlsx(new Uint8Array(readFileSync(xlsxPath)), basename(xlsxPath))
  const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, { useColumnIndex: true })
  mkdirSync(dirname(workpaperPath), { recursive: true })
  writeFileAtomically(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true })))

  return createFileBackedWorkPaperMcpToolServerFromFile({
    workpaperPath,
    writable: input.writable ?? false,
  })
}

function writeFileAtomically(path: string, contents: string): void {
  const tempPath = resolve(dirname(path), `.${basename(path)}.${process.pid.toString()}.tmp`)
  writeFileSync(tempPath, contents)
  renameSync(tempPath, path)
}

export { createFileBackedWorkPaperMcpToolServerFromXlsxFile, type FileBackedWorkPaperMcpFromXlsxOptions }
