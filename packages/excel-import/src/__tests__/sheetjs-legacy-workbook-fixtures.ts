import * as XLSX from 'xlsx'

export type SheetJsFallbackWorkbook = ReturnType<typeof XLSX.read>

export function readSheetJsFallbackWorkbook(bytes: Uint8Array): SheetJsFallbackWorkbook {
  return XLSX.read(bytes, {
    type: 'array',
    bookFiles: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: false,
  })
}

export function buildBinaryWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Name', 'Value'],
    ['alpha', 12],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['notes']]), 'Sheet2')
  return XLSX.write(workbook, { bookType: 'xlsb', type: 'buffer' })
}

export function buildLegacyWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Department', 'Amount'],
    ['Operations', 1250],
    ['Finance', 1800],
  ])
  sheet.C2 = { t: 'n', f: 'B2+B3', v: 3050 }
  sheet['!ref'] = 'A1:C3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Salary')
  return XLSX.write(workbook, { bookType: 'xls', type: 'buffer' })
}

export function buildNamespacedFormulaWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[1], [2]])
  sheet.A3 = { t: 'n', f: 'msoxl:=SUM(A1:A2)', v: 3 }
  sheet.B3 = { t: 'n', f: 'of:=SUM(A1:A2)', v: 3 }
  sheet['!ref'] = 'A1:B3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Expenses')
  return XLSX.write(workbook, { bookType: 'ods', type: 'buffer' })
}
