import { mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inflateRawSync } from 'node:zlib'

import { describe, expect, it } from 'vitest'

import {
  exportXlsxSourceLiteralPatches,
  exportXlsxSourceLiteralPatchesToFileAsync,
  getZipText,
  readLazyXlsxZipEntryCompressedSource,
  readXlsxZipEntries,
  readXlsxZipEntriesLazy,
  zipSourcePreservingEntries,
  type XlsxSourceReader,
} from '../index.js'

function minimalWorkbookBytes(extraEntries: Record<string, Uint8Array> = {}): Uint8Array {
  return zipSourcePreservingEntries({
    '[Content_Types].xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`),
    '_rels/.rels': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/workbook.xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Revenue &amp; Ops" sheetId="1" r:id="rId1"/>
  </sheets>
  <calcPr calcMode="manual"/>
</workbook>`),
    'xl/_rels/workbook.xml.rels': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>old</t></is></c></row>
  </sheetData>
</worksheet>`),
    'xl/calcChain.xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>`),
    ...extraEntries,
  })
}

function formulaCacheWorkbookBytes(): Uint8Array {
  return minimalWorkbookBytes({
    'xl/worksheets/sheet1.xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C1"/>
  <sheetData>
    <row r="1">
      <c r="A1"><v>4</v></c>
      <c r="B1"><f>A1*2</f><v>8</v></c>
      <c r="C1" t="str"><f>A1&amp;" units"</f><v>old units</v></c>
    </row>
  </sheetData>
</worksheet>`),
  })
}

function splitSharedFormulaCacheWorkbookBytes(): Uint8Array {
  const padding = ' '.repeat(65_512)
  return minimalWorkbookBytes({
    'xl/worksheets/sheet1.xml': new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A2"/>
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row>
    ${padding}<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function guardedUntouchedEntrySource(): {
  readonly source: XlsxSourceReader
  readonly untouchedBytes: Uint8Array
  readonly untouchedPath: string
  readonly untouchedInflateCount: () => number
} {
  const untouchedBytes = new Uint8Array(256 * 1024)
  for (let index = 0; index < untouchedBytes.byteLength; index += 1) {
    untouchedBytes[index] = index & 0xff
  }
  const untouchedPath = 'xl/media/untouched.bin'
  const sourceBytes = minimalWorkbookBytes({ [untouchedPath]: untouchedBytes })
  const untouchedSourceEntry = readLazyXlsxZipEntryCompressedSource(readXlsxZipEntriesLazy(sourceBytes), untouchedPath)
  if (!untouchedSourceEntry) {
    throw new Error('Expected untouched entry to have lazy compressed source metadata')
  }

  let untouchedInflateCount = 0
  const assertNotUntouchedInflate = (start: number, end: number): void => {
    if (start === untouchedSourceEntry.dataStart && end === untouchedSourceEntry.dataEnd) {
      untouchedInflateCount += 1
      throw new Error('untouched entry should be copied from compressed source')
    }
  }
  return {
    source: {
      byteLength: sourceBytes.byteLength,
      readBytes() {
        throw new Error('readBytes should not be called')
      },
      readRange(start, end) {
        return sourceBytes.subarray(start, end)
      },
      inflateRawRange(start, end) {
        assertNotUntouchedInflate(start, end)
        return inflateRawSync(sourceBytes.subarray(start, end))
      },
      async inflateRawRangeChunksAsync(start, end) {
        assertNotUntouchedInflate(start, end)
        return false
      },
    },
    untouchedBytes,
    untouchedPath,
    untouchedInflateCount: () => untouchedInflateCount,
  }
}

function guardedPatchedWorksheetSource(): {
  readonly source: XlsxSourceReader
  readonly fullWorksheetInflateCount: () => number
} {
  const rowCount = 5_000
  const worksheetXml = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A${String(rowCount)}"/>
  <sheetData>
    ${Array.from({ length: rowCount }, (_entry, index) => {
      const row = index + 1
      return `<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`
    }).join('')}
  </sheetData>
</worksheet>`)
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sourceBytes = minimalWorkbookBytes({ [sheetPath]: worksheetXml })
  const sheetSourceEntry = readLazyXlsxZipEntryCompressedSource(readXlsxZipEntriesLazy(sourceBytes), sheetPath)
  if (!sheetSourceEntry) {
    throw new Error('Expected worksheet entry to have lazy compressed source metadata')
  }

  let fullWorksheetInflateCount = 0
  const assertNotFullWorksheetInflate = (start: number, end: number): void => {
    if (start === sheetSourceEntry.dataStart && end === sheetSourceEntry.dataEnd) {
      fullWorksheetInflateCount += 1
      throw new Error('patched worksheet should be streamed instead of fully inflated')
    }
  }
  return {
    source: {
      byteLength: sourceBytes.byteLength,
      readBytes() {
        throw new Error('readBytes should not be called')
      },
      readRange(start, end) {
        return sourceBytes.subarray(start, end)
      },
      readRangeInto(start, end, target) {
        const chunk = sourceBytes.subarray(start, end)
        target.set(chunk)
        return target.subarray(0, chunk.byteLength)
      },
      inflateRawRange(start, end) {
        assertNotFullWorksheetInflate(start, end)
        return inflateRawSync(sourceBytes.subarray(start, end))
      },
    },
    fullWorksheetInflateCount: () => fullWorksheetInflateCount,
  }
}

describe('@bilig/xlsx source-preserving literal patches', () => {
  it('patches scalar cells, inserts missing cells, and invalidates calcChain without SheetJS', () => {
    const exported = exportXlsxSourceLiteralPatches({
      source: minimalWorkbookBytes(),
      sheetNames: ['Revenue & Ops'],
      patches: [
        { sheetName: 'Revenue & Ops', address: 'A1', value: 42 },
        { sheetName: 'Revenue & Ops', address: 'B1', value: 'new & checked' },
        { sheetName: 'Revenue & Ops', address: 'C3', value: true },
      ],
    })

    const zip = readXlsxZipEntries(exported)
    const sheetXml = getZipText(zip, 'xl/worksheets/sheet1.xml')
    const workbookXml = getZipText(zip, 'xl/workbook.xml')
    const workbookRelationshipsXml = getZipText(zip, 'xl/_rels/workbook.xml.rels')
    const contentTypesXml = getZipText(zip, '[Content_Types].xml')

    expect(sheetXml).toContain('<dimension ref="A1:C3"/>')
    expect(sheetXml).toContain('<c r="A1"><v>42</v></c>')
    expect(sheetXml).toContain('<c r="B1" t="inlineStr"><is><t>new &amp; checked</t></is></c>')
    expect(sheetXml).toContain('<c r="C3" t="b"><v>1</v></c>')
    expect(workbookXml).toContain('calcMode="auto"')
    expect(workbookXml).toContain('fullCalcOnLoad="1"')
    expect(workbookXml).toContain('forceFullCalc="1"')
    expect(zip['xl/calcChain.xml']).toBeUndefined()
    expect(workbookRelationshipsXml).not.toContain('calcChain')
    expect(contentTypesXml).not.toContain('calcChain.xml')
  })

  it('patches cached formula results without replacing formula XML', () => {
    const exported = exportXlsxSourceLiteralPatches({
      source: formulaCacheWorkbookBytes(),
      sheetNames: ['Revenue & Ops'],
      patches: [
        { sheetName: 'Revenue & Ops', address: 'A1', value: 10 },
        { sheetName: 'Revenue & Ops', address: 'B1', value: 20, preserveFormula: true },
        { sheetName: 'Revenue & Ops', address: 'C1', value: '10 & checked', preserveFormula: true },
      ],
    })

    const sheetXml = getZipText(readXlsxZipEntries(exported), 'xl/worksheets/sheet1.xml')

    expect(sheetXml).toContain('<c r="A1"><v>10</v></c>')
    expect(sheetXml).toContain('<c r="B1"><f>A1*2</f><v>20</v></c>')
    expect(sheetXml).toContain('<c r="C1" t="str"><f>A1&amp;" units"</f><v>10 &amp; checked</v></c>')
    expect(sheetXml).not.toContain('inlineStr"><f>')
  })

  it('writes file-backed patches without reading the whole source through readBytes', async () => {
    const sourceBytes = minimalWorkbookBytes()
    let readBytesCalled = false
    const source: XlsxSourceReader = {
      byteLength: sourceBytes.byteLength,
      readBytes() {
        readBytesCalled = true
        throw new Error('readBytes should not be called')
      },
      readRange(start, end) {
        return sourceBytes.subarray(start, end)
      },
    }
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-patch-'))
    const outputPath = join(tempDir, 'patched.xlsx')
    try {
      const result = await exportXlsxSourceLiteralPatchesToFileAsync({
        source,
        outputPath,
        sheetNames: ['Revenue & Ops'],
        patches: [{ sheetName: 'Revenue & Ops', address: 'B1', value: 'file-backed' }],
      })

      expect(readBytesCalled).toBe(false)
      expect(result.bytesWritten).toBe(statSync(outputPath).size)
      const zip = readXlsxZipEntries(new Uint8Array(readFileSync(outputPath)))
      expect(getZipText(zip, 'xl/worksheets/sheet1.xml')).toContain('file-backed')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('copies untouched lazy entries from their compressed source without inflating them', () => {
    const { source, untouchedBytes, untouchedPath, untouchedInflateCount } = guardedUntouchedEntrySource()

    const exported = exportXlsxSourceLiteralPatches({
      source,
      sheetNames: ['Revenue & Ops'],
      patches: [{ sheetName: 'Revenue & Ops', address: 'B1', value: 'compressed-copy' }],
    })

    expect(untouchedInflateCount()).toBe(0)
    const zip = readXlsxZipEntries(exported)
    expect(Buffer.from(zip[untouchedPath] ?? new Uint8Array()).equals(Buffer.from(untouchedBytes))).toBe(true)
    expect(getZipText(zip, 'xl/worksheets/sheet1.xml')).toContain('compressed-copy')
  })

  it('streams sync patched worksheet entries without fully inflating the worksheet', () => {
    const { source, fullWorksheetInflateCount } = guardedPatchedWorksheetSource()

    const exported = exportXlsxSourceLiteralPatches({
      source,
      sheetNames: ['Revenue & Ops'],
      patches: [{ sheetName: 'Revenue & Ops', address: 'A2500', value: 99_001 }],
    })

    expect(fullWorksheetInflateCount()).toBe(0)
    const zip = readXlsxZipEntries(exported)
    expect(getZipText(zip, 'xl/worksheets/sheet1.xml')).toContain('<c r="A2500"><v>99001</v></c>')
  })

  it('file-backed export copies untouched lazy entries from compressed source without inflating them', async () => {
    const { source, untouchedBytes, untouchedPath, untouchedInflateCount } = guardedUntouchedEntrySource()
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-patch-'))
    const outputPath = join(tempDir, 'patched.xlsx')
    try {
      await exportXlsxSourceLiteralPatchesToFileAsync({
        source,
        outputPath,
        sheetNames: ['Revenue & Ops'],
        patches: [{ sheetName: 'Revenue & Ops', address: 'B1', value: 'file-compressed-copy' }],
      })

      expect(untouchedInflateCount()).toBe(0)
      const zip = readXlsxZipEntries(new Uint8Array(readFileSync(outputPath)))
      expect(Buffer.from(zip[untouchedPath] ?? new Uint8Array()).equals(Buffer.from(untouchedBytes))).toBe(true)
      expect(getZipText(zip, 'xl/worksheets/sheet1.xml')).toContain('file-compressed-copy')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('file-backed export patches shared formula cache cells split across stream chunks', async () => {
    const sourceBytes = splitSharedFormulaCacheWorkbookBytes()
    let readBytesCalled = false
    const source: XlsxSourceReader = {
      byteLength: sourceBytes.byteLength,
      readBytes() {
        readBytesCalled = true
        throw new Error('readBytes should not be called')
      },
      readRange(start, end) {
        return sourceBytes.subarray(start, end)
      },
      readRangeInto(start, end, target) {
        const chunk = sourceBytes.subarray(start, end)
        target.set(chunk)
        return target.subarray(0, chunk.byteLength)
      },
    }
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-patch-'))
    const outputPath = join(tempDir, 'patched.xlsx')
    try {
      await exportXlsxSourceLiteralPatchesToFileAsync({
        source,
        outputPath,
        sheetNames: ['Revenue & Ops'],
        patches: [{ sheetName: 'Revenue & Ops', address: 'A2', value: 3, preserveFormula: true }],
      })

      expect(readBytesCalled).toBe(false)
      const sheetXml = getZipText(readXlsxZipEntries(new Uint8Array(readFileSync(outputPath))), 'xl/worksheets/sheet1.xml')
      expect(sheetXml).toContain('<c r="A2"><f t="shared" si="0"/><v>3</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})
