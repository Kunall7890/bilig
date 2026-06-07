import { unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { readSheetJsFallbackWorkbook } from './sheetjs-legacy-workbook-fixtures.js'
import { readImportedWorkbookConditionalFormats } from '../xlsx-conditional-formats.js'
import { readImportedWorkbookDataModelArtifacts } from '../xlsx-data-model-artifacts.js'
import { readImportedWorkbookFileStyles, readImportedWorkbookSheetDimensions, readImportedWorkbookStyleArtifacts } from '../xlsx-styles.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('XLSX fallback lazy artifact readers', () => {
  it('preserves lazy ZIP entries while reading SheetJS fallback metadata artifacts', () => {
    const bytes = buildWorkbookWithUnrelatedHeavyPart()
    const workbook = readSheetJsFallbackWorkbook(bytes)
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
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Data',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Team' },
            { address: 'B1', row: 0, col: 1, value: 'Score' },
            { address: 'A2', row: 1, col: 0, value: 'Finance' },
            { address: 'B2', row: 1, col: 1, value: 7 },
            { address: 'A3', row: 2, col: 0, value: 'Ops' },
            { address: 'B3', row: 2, col: 1, value: 3 },
          ],
        },
      ],
    }),
  )
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
