import { unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { readImportedWorkbookConditionalFormats } from '../xlsx-conditional-formats.js'
import { readImportedWorkbookDataModelArtifacts } from '../xlsx-data-model-artifacts.js'
import { readImportedWorkbookFileStyles, readImportedWorkbookSheetDimensions, readImportedWorkbookStyleArtifacts } from '../xlsx-styles.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('XLSX fallback lazy artifact readers', () => {
  it('preserves lazy ZIP entries while reading SheetJS fallback metadata artifacts', () => {
    const bytes = buildWorkbookWithUnrelatedHeavyPart()
    const workbook = XLSX.read(bytes, {
      type: 'array',
      bookFiles: true,
      cellFormula: true,
      cellNF: true,
      cellStyles: false,
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/model/item.data', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('unrelated data model bytes should stay lazy during fallback artifact reads')
      },
    })

    expect(() => {
      readImportedWorkbookSheetDimensions(workbook, workbook.SheetNames, zip)
      readImportedWorkbookFileStyles(workbook, workbook.SheetNames, {}, zip)
      readImportedWorkbookStyleArtifacts(workbook, workbook.SheetNames, zip)
      readImportedWorkbookConditionalFormats(zip, workbook.SheetNames)
    }).not.toThrow()

    const dataModelArtifacts = readImportedWorkbookDataModelArtifacts(zip)
    expect(dataModelArtifacts?.parts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ path: 'xl/model/item.data', byteLength: 1024 * 1024 }),
        expect.objectContaining({ path: 'xl/queries/query1.xml', byteLength: 512 * 1024 }),
      ]),
    )

    const descriptor = Object.getOwnPropertyDescriptor(zip, 'xl/model/item.data')
    expect(descriptor && 'get' in descriptor && typeof descriptor.get).toBe('function')
  })
})

function buildWorkbookWithUnrelatedHeavyPart(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Team', 'Score'],
    ['Finance', 7],
    ['Ops', 3],
  ])
  sheet['!ref'] = 'A1:B3'
  XLSX.utils.book_append_sheet(workbook, sheet, 'Data')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(
    new TextDecoder()
      .decode(zip['xl/worksheets/sheet1.xml'])
      .replace(
        '</worksheet>',
        '<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting></worksheet>',
      ),
  )
  zip['xl/model/item.data'] = new Uint8Array(1024 * 1024)
  zip['xl/queries/query1.xml'] = new Uint8Array(512 * 1024)
  return zipSync(zip)
}
