import * as XLSX from 'xlsx'

import { parseCsv, parseCsvCellInput, resolveCsvParseOptions, type CsvParseOptions } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { CSV_CONTENT_TYPE } from './workbook-import-content-types.js'
import { createSheetPreview, normalizeCsvSheetName, normalizeWorkbookName } from './workbook-import-helpers.js'
import { createWorkbookPreview } from './workbook-import-preview.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import {
  attachImportedRuntimeCoordinates,
  createImportedRuntimeSheetCells,
  pushImportedSnapshotCell,
  type ImportedRuntimeCellCoordinate,
} from './xlsx-import-runtime-coordinates.js'

export type CsvImportOptions = CsvParseOptions

export function importCsv(text: string, fileName: string, options: CsvImportOptions = {}): ImportedWorkbook {
  const workbookName = normalizeWorkbookName(fileName)
  const sheetName = normalizeCsvSheetName(workbookName)
  const csvOptions = resolveCsvParseOptions(text, options)
  const rows = parseCsv(text, csvOptions)
  const textColumnIndexes = inferCsvTextColumnIndexes(rows)
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  const runtimeCellCoords: ImportedRuntimeCellCoordinate[] = []
  let nonEmptyCellCount = 0
  let hasRaggedRows = false
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0)

  rows.forEach((row, rowIndex) => {
    if (row.length !== columnCount) {
      hasRaggedRows = true
    }
    row.forEach((raw, colIndex) => {
      const parsed =
        textColumnIndexes.has(colIndex) && rowIndex > 0 && raw.trim() !== '' ? { value: raw } : parseCsvCellInput(raw, csvOptions)
      if (!parsed) {
        return
      }
      nonEmptyCellCount += 1
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
      if (parsed.formula !== undefined) {
        pushImportedSnapshotCell(cells, runtimeCellCoords, { address, formula: parsed.formula }, rowIndex, colIndex)
        return
      }
      if (parsed.value !== undefined) {
        pushImportedSnapshotCell(cells, runtimeCellCoords, { address, value: parsed.value }, rowIndex, colIndex)
      }
    })
  })

  const warnings = hasRaggedRows ? ['CSV rows had inconsistent column counts. Missing cells were treated as blanks.'] : []
  const previewSheet = createSheetPreview({
    name: sheetName,
    rowCount: rows.length,
    columnCount,
    nonEmptyCellCount,
    readCellText: (row, col) => rows[row]?.[col] ?? '',
  })

  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: workbookName,
    },
    sheets: [
      {
        id: 1,
        name: sheetName,
        order: 0,
        cells,
      },
    ],
  }
  const runtimeSheetCells = [
    createImportedRuntimeSheetCells({
      sheetName,
      coords: runtimeCellCoords,
      width: columnCount,
      height: rows.length,
    }),
  ]

  return {
    snapshot: attachImportedRuntimeCoordinates(snapshot, runtimeSheetCells),
    workbookName,
    sheetNames: [sheetName],
    warnings,
    preview: createWorkbookPreview({
      contentType: CSV_CONTENT_TYPE,
      fileName,
      fileSizeBytes: new TextEncoder().encode(text).byteLength,
      workbookName,
      sheets: [previewSheet],
      warnings,
    }),
  }
}

function inferCsvTextColumnIndexes(rows: readonly (readonly string[])[]): Set<number> {
  const header = rows[0]
  const textColumnIndexes = new Set<number>()
  if (!header) {
    return textColumnIndexes
  }

  header.forEach((rawHeader, colIndex) => {
    const headerText = rawHeader.trim().toLowerCase().replaceAll('_', ' ').replaceAll('-', ' ')
    if (isIdentifierLikeCsvHeader(headerText)) {
      textColumnIndexes.add(colIndex)
    }
  })
  return textColumnIndexes
}

function isIdentifierLikeCsvHeader(headerText: string): boolean {
  return /^(?:account|acct|id|code|sku)(?: (?:id|number|no|code))?$/u.test(headerText)
}
