import { encodeCellAddress, writeSimpleXlsxWorkbook, type SimpleXlsxCell } from '@bilig/xlsx'

export type BiligXlsxFixtureCellValue = boolean | number | string | null

export interface BiligXlsxFixtureWorkbookInput {
  readonly sheetName: string
  readonly rows: ReadonlyArray<ReadonlyArray<BiligXlsxFixtureCellValue>>
  readonly formulaStrings?: 'as-formulas' | 'as-blanks'
}

export function writeBiligXlsxFixtureWorkbook(input: BiligXlsxFixtureWorkbookInput): Uint8Array {
  const formulaStrings = input.formulaStrings ?? 'as-formulas'
  const cells: SimpleXlsxCell[] = []
  for (let rowIndex = 0; rowIndex < input.rows.length; rowIndex += 1) {
    const row = input.rows[rowIndex] ?? []
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const value = row[colIndex]
      if (value === null) {
        continue
      }
      const address = encodeCellAddress({ r: rowIndex, c: colIndex })
      if (typeof value === 'string' && value.startsWith('=')) {
        if (formulaStrings === 'as-formulas') {
          cells.push({ address, row: rowIndex, col: colIndex, formula: value.slice(1) })
        }
        continue
      }
      cells.push({ address, row: rowIndex, col: colIndex, value })
    }
  }
  return writeSimpleXlsxWorkbook({ sheets: [{ name: input.sheetName, cells }] })
}
