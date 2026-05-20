import type * as XLSX from 'xlsx'

import type { PreservedVbaProjectCodeNames } from './xlsx-macros.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

export function readImportedMacroCodeNames(workbook: XLSX.WorkBook): PreservedVbaProjectCodeNames {
  const workbookMetadata = isRecord(workbook.Workbook) ? workbook.Workbook : undefined
  const workbookProperties = isRecord(workbookMetadata?.['WBProps']) ? workbookMetadata['WBProps'] : undefined
  const workbookCodeName = readNonEmptyString(workbookProperties?.['CodeName'])
  const workbookSheets = workbookMetadata?.['Sheets']
  const sheetCodeNames = Array.isArray(workbookSheets)
    ? workbookSheets.flatMap((entry) => {
        if (!isRecord(entry)) {
          return []
        }
        const sheetName = readNonEmptyString(entry['name'])
        const codeName = readNonEmptyString(entry['CodeName'])
        return sheetName && codeName ? [{ sheetName, codeName }] : []
      })
    : []
  return {
    ...(workbookCodeName ? { workbookCodeName } : {}),
    ...(sheetCodeNames.length > 0 ? { sheetCodeNames } : {}),
  }
}
