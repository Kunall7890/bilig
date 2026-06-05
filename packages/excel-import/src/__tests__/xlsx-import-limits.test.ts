import type { Unzipped } from 'fflate'
import { describe, expect, it } from 'vitest'

import { planXlsxImportRoute } from '../xlsx-import-limits.js'

describe('xlsx import route planning', () => {
  it('keeps small default imports on the full-fidelity fallback threshold', () => {
    const route = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 8_000,
      options: {},
      inspection: null,
    })

    expect(route.createLargeSimpleImportOptions().minByteLength).toBeUndefined()
  })

  it('allows formula recalc callers to opt into small native simple imports', () => {
    const route = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 8_000,
      options: { preferNativeSimpleImport: true },
      inspection: null,
    })

    expect(route.createLargeSimpleImportOptions().minByteLength).toBe(0)
  })
})

function zipWith(paths: readonly string[]): Unzipped {
  return Object.fromEntries(paths.map((path) => [path, new Uint8Array()]))
}
