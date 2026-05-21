import { describe, expect, it } from 'vitest'

import { readLargeSimpleWorkbookNumberFormatsFromChunks } from '../xlsx-large-simple-number-formats.js'
import { readLargeSimpleWorkbookStyleArtifactsFromChunks, readLargeSimpleWorkbookStylesFromChunks } from '../xlsx-large-simple-styles.js'

const encoder = new TextEncoder()

describe('large simple styles streaming', () => {
  it('discards unneeded indexed style children while waiting for the closing tag', () => {
    const retainedBufferLengths: number[] = []
    const largeUnneededFillPayload = 'x'.repeat(100_000)
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="0" fillId="1" fontId="1" applyFill="1" applyFont="1"/></cellXfs>',
      '<fills count="2"><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill><unused>',
      largeUnneededFillPayload.slice(0, 40_000),
      largeUnneededFillPayload.slice(40_000, 80_000),
      largeUnneededFillPayload.slice(80_000),
      '</unused></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>',
      '<fonts count="2"><font/><font><b/><name val="Inter"/><sz val="11"/></font></fonts>',
      '</styleSheet>',
    ]

    const styles = readLargeSimpleWorkbookStylesFromChunks(
      (onChunk) => {
        for (const chunk of chunks) {
          onChunk(encoder.encode(chunk))
        }
        return true
      },
      new Set([1]),
      {
        onRetainedBufferLength: (length) => retainedBufferLengths.push(length),
      },
    )

    expect(styles?.get(1)).toEqual({
      fill: { backgroundColor: '#ffcc00' },
      font: { bold: true, family: 'Inter', size: 11 },
    })
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('streams number formats without dropping the visual style on the same xf', () => {
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="164" fillId="1" fontId="0" applyFill="1" applyNumberFormat="1"/></cellXfs>',
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>',
      '<fonts count="1"><font/></fonts>',
      '</styleSheet>',
    ]
    const readChunks = (onChunk: (chunk: Uint8Array) => void): boolean => {
      for (const chunk of chunks) {
        onChunk(encoder.encode(chunk))
      }
      return true
    }

    expect(readLargeSimpleWorkbookStylesFromChunks(readChunks, new Set([1]))?.get(1)).toEqual({
      fill: { backgroundColor: '#ffcc00' },
    })
    expect(readLargeSimpleWorkbookNumberFormatsFromChunks(readChunks, new Set([1]))?.get(1)).toBe('00000')
  })

  it('collects style and number-format artifacts with a shared component pass', () => {
    let readPassCount = 0
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>',
      '<fonts count="2"><font/><font><b/><name val="Inter"/><sz val="11"/></font></fonts>',
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill></fills>',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="164" fillId="1" fontId="1" applyFill="1" applyFont="1" applyNumberFormat="1"/></cellXfs>',
      '</styleSheet>',
    ]
    const readChunks = (onChunk: (chunk: Uint8Array) => void): boolean => {
      readPassCount += 1
      for (const chunk of chunks) {
        onChunk(encoder.encode(chunk))
      }
      return true
    }

    const artifacts = readLargeSimpleWorkbookStyleArtifactsFromChunks(readChunks, new Set([1]))

    expect(artifacts.stylesByIndex?.get(1)).toEqual({
      fill: { backgroundColor: '#ffcc00' },
      font: { bold: true, family: 'Inter', size: 11 },
    })
    expect(artifacts.numberFormatsByStyleIndex?.get(1)).toBe('00000')
    expect(readPassCount).toBe(2)
  })

  it('keeps number formats when an unsupported visual style forces style fallback', () => {
    let readPassCount = 0
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>',
      '<borders count="2"><border/><border><left style="thin"/></border></borders>',
      '<cellXfs count="2"><xf numFmtId="0" fillId="0" fontId="0"/><xf numFmtId="164" borderId="1" applyBorder="1" applyNumberFormat="1"/></cellXfs>',
      '</styleSheet>',
    ]
    const readChunks = (onChunk: (chunk: Uint8Array) => void): boolean => {
      readPassCount += 1
      for (const chunk of chunks) {
        onChunk(encoder.encode(chunk))
      }
      return true
    }

    const artifacts = readLargeSimpleWorkbookStyleArtifactsFromChunks(readChunks, new Set([1]))

    expect(artifacts.stylesByIndex).toBeNull()
    expect(artifacts.numberFormatsByStyleIndex?.get(1)).toBe('00000')
    expect(readPassCount).toBe(2)
  })
})
