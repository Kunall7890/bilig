import type * as XLSX from 'xlsx'

import type { WorkbookSheetVisibilitySnapshot, WorkbookSnapshot } from '@bilig/protocol'

function importedSheetVisibility(value: unknown): WorkbookSheetVisibilitySnapshot | undefined {
  if (value === 1 || value === '1') {
    return 'hidden'
  }
  if (value === 2 || value === '2') {
    return 'veryHidden'
  }
  return undefined
}

function exportedSheetHidden(value: WorkbookSheetVisibilitySnapshot | undefined): 1 | 2 | undefined {
  if (value === 'hidden') {
    return 1
  }
  if (value === 'veryHidden') {
    return 2
  }
  return undefined
}

export function readImportedWorkbookSheetVisibilities(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetVisibilitySnapshot> {
  const sheetProps = workbook.Workbook?.Sheets ?? []
  const visibilitiesBySheet = new Map<string, WorkbookSheetVisibilitySnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const visibility = importedSheetVisibility(sheetProps[sheetIndex]?.Hidden)
    if (visibility) {
      visibilitiesBySheet.set(sheetName, visibility)
    }
  })

  return visibilitiesBySheet
}

export function applyExportSheetVisibilitiesToWorkbook(workbook: XLSX.WorkBook, snapshot: WorkbookSnapshot): void {
  const orderedSheets = snapshot.sheets.toSorted((left, right) => left.order - right.order)
  if (!orderedSheets.some((sheet) => sheet.metadata?.visibility)) {
    return
  }

  const existingSheets = workbook.Workbook?.Sheets ?? []
  const sheets = workbook.SheetNames.map((sheetName, sheetIndex) => {
    const hidden = exportedSheetHidden(orderedSheets[sheetIndex]?.metadata?.visibility)
    const sheet = {
      ...existingSheets[sheetIndex],
      name: sheetName,
    }
    delete sheet.Hidden
    if (hidden !== undefined) {
      sheet.Hidden = hidden
    }
    return sheet
  })

  workbook.Workbook = {
    ...workbook.Workbook,
    Sheets: sheets,
  }
}
