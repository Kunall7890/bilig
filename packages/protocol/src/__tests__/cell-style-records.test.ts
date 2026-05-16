import { describe, expect, it } from 'vitest'
import { sanitizeCellStyleRecord } from '../cell-style-records.js'

describe('sanitizeCellStyleRecord', () => {
  it('drops malformed fields and preserves valid protocol style fields', () => {
    expect(
      sanitizeCellStyleRecord('style-1', {
        id: 'ignored',
        fill: { backgroundColor: '#ffee00' },
        font: {
          family: 'Inter',
          size: 13,
          bold: true,
          underline: false,
          color: '#111111',
          shadow: true,
        },
        alignment: {
          horizontal: 'right',
          vertical: 'middle',
          indent: 2,
          textRotation: 45,
          wrap: true,
          readingOrder: Number.POSITIVE_INFINITY,
          fake: 'ignored',
        },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#333333' },
          bottom: { style: 'wave', weight: 'thin', color: '#333333' },
        },
        protection: { locked: true, hidden: 0 },
        arbitrary: { trusted: false },
      }),
    ).toEqual({
      id: 'style-1',
      fill: { backgroundColor: '#ffee00' },
      font: {
        family: 'Inter',
        size: 13,
        bold: true,
        underline: false,
        color: '#111111',
      },
      alignment: {
        horizontal: 'right',
        vertical: 'middle',
        indent: 2,
        textRotation: 45,
        wrap: true,
      },
      borders: {
        top: { style: 'solid', weight: 'thin', color: '#333333' },
      },
      protection: { locked: true },
    })
  })

  it('rejects missing ids and non-object payloads', () => {
    expect(sanitizeCellStyleRecord('', { fill: { backgroundColor: '#ffee00' } })).toBeNull()
    expect(sanitizeCellStyleRecord('style-1', null)).toBeNull()
  })
})
