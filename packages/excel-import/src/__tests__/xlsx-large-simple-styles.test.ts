import { describe, expect, it } from 'vitest'

import { readLargeSimpleWorkbookNumberFormatsFromChunks } from '../xlsx-large-simple-number-formats.js'
import {
  hasUnsupportedLargeSimpleWorkbookStylesFromChunks,
  inspectLargeSimpleWorkbookStyleSupportFromChunks,
  readLargeSimpleWorkbookStyleArtifactsFromChunks,
  readLargeSimpleWorkbookStylesFromChunks,
} from '../xlsx-large-simple-styles.js'

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

  it('detects unsupported required cellXfs without retaining unneeded style XML', () => {
    const retainedBufferLengths: number[] = []
    const largeUnneededAlignment = 'x'.repeat(100_000)
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="3">',
      '<xf numFmtId="0" fillId="0" fontId="0"><alignment textRotation="0">',
      largeUnneededAlignment.slice(0, 50_000),
      largeUnneededAlignment.slice(50_000),
      '</alignment></xf>',
      '<xf numFmtId="0" borderId="1" applyBorder="1"/>',
      '<xf numFmtId="0" fillId="0" fontId="0"/>',
      '</cellXfs></styleSheet>',
    ]

    const hasUnsupportedStyles = hasUnsupportedLargeSimpleWorkbookStylesFromChunks(
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

    expect(hasUnsupportedStyles).toBe(true)
    expect(retainedBufferLengths.length).toBeGreaterThan(0)
    expect(Math.max(...retainedBufferLengths)).toBeLessThan(1024)
  })

  it('stops style support scanning after a required unsupported xf is found', () => {
    let chunksRead = 0
    const hasUnsupportedStyles = hasUnsupportedLargeSimpleWorkbookStylesFromChunks(
      (onChunk) => {
        for (const chunk of [
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="3">',
          '<xf numFmtId="0" borderId="1" applyBorder="1"/>',
          '<xf numFmtId="0" fillId="0" applyFill="1"/><xf numFmtId="0"/></cellXfs></styleSheet>',
        ]) {
          chunksRead += 1
          if (onChunk(encoder.encode(chunk)) === false) {
            break
          }
        }
        return true
      },
      new Set([0]),
    )

    expect(hasUnsupportedStyles).toBe(true)
    expect(chunksRead).toBe(2)
  })

  it('detects potential visual style ranges from required cellXfs without number-format false positives', () => {
    const chunks = [
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="3">',
      '<xf numFmtId="164" applyNumberFormat="1"/>',
      '<xf numFmtId="0" fillId="1" applyFill="1"/>',
      '<xf numFmtId="0" applyAlignment="1"><alignment horizontal="center"/></xf>',
      '</cellXfs></styleSheet>',
    ]
    const readChunks = (onChunk: (chunk: Uint8Array) => void): boolean => {
      for (const chunk of chunks) {
        onChunk(encoder.encode(chunk))
      }
      return true
    }

    expect(inspectLargeSimpleWorkbookStyleSupportFromChunks(readChunks, new Set([0]))).toEqual({
      hasUnsupportedStyles: false,
      hasPotentialVisualStyles: false,
    })
    expect(inspectLargeSimpleWorkbookStyleSupportFromChunks(readChunks, new Set([1, 2]))).toEqual({
      hasUnsupportedStyles: false,
      hasPotentialVisualStyles: true,
    })
  })

  it('returns null when a required cellXfs entry is missing', () => {
    const hasUnsupportedStyles = hasUnsupportedLargeSimpleWorkbookStylesFromChunks(
      (onChunk) => {
        onChunk(encoder.encode('<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>'))
        return true
      },
      new Set([2]),
    )

    expect(hasUnsupportedStyles).toBeNull()
  })
})
