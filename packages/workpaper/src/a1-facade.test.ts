import { describe, expect, it } from 'vitest'

import { buildA1WorkPaper } from './index.js'

describe('@bilig/workpaper A1 facade export', () => {
  it('exposes the A1 WorkPaper facade through the scoped package entrypoint', () => {
    const book = buildA1WorkPaper({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })

    try {
      const proof = book.setCellAndReadback('Inputs!B2', 48, {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(true)
      expect(proof.afterReadback.displayValues).toEqual([['57600']])
      expect(proof.restoredReadback.displayValues).toEqual([['57600']])
    } finally {
      book.dispose()
    }
  })
})
