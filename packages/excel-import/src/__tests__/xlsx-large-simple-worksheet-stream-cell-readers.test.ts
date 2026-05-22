import { describe, expect, it } from 'vitest'

import { readInlineStringCellValue, readRichTextCellArtifact } from '../xlsx-large-simple-worksheet-stream-cell-readers.js'

const encoder = new TextEncoder()

describe('large simple worksheet stream cell readers', () => {
  it('reads empty inline strings without requiring rich-text XML materialization', () => {
    expect(readInlineString('<is><t></t></is>')).toBe('')
    expect(readInlineString('<is><t/></is>')).toBe('')
    expect(readInlineString('<is/>')).toBe('')
  })

  it('matches worksheet text normalization for simple and rich inline strings', () => {
    expect(readInlineString('<is><t>Food &amp; tobacco</t></is>')).toBe('Food & tobacco')
    expect(readInlineString('<is><r><t>A</t></r><r><t>1</t></r></is>')).toBe('A1')
    expect(readInlineString('<is><t>_x005F_</t></is>')).toBe('_')
  })

  it('detects rich inline text only when rich runs are present', () => {
    expect(readInlineRichText('<is><t>plain</t></is>')).toBeUndefined()
    expect(readInlineRichText('<is><r><t>rich</t></r></is>')).toEqual({
      address: 'A1',
      text: 'rich',
      storage: 'inlineString',
      xml: '<is><r><t>rich</t></r></is>',
    })
  })
})

function readInlineString(xml: string): string | undefined {
  const bytes = encoder.encode(`<c t="inlineStr">${xml}</c>`)
  return readInlineStringCellValue(bytes, '<c t="inlineStr">'.length, bytes.length - '</c>'.length)
}

function readInlineRichText(xml: string): ReturnType<typeof readRichTextCellArtifact> {
  const bytes = encoder.encode(`<c t="inlineStr">${xml}</c>`)
  return readRichTextCellArtifact(bytes, '<c t="inlineStr">'.length, bytes.length - '</c>'.length, 0, 0, 'inlineStr', null, [])
}
