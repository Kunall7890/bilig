import { describe, expect, it } from 'vitest'

import { buildA1WorkPaper, restoreA1WorkPaper } from '../index.js'

describe('A1 WorkPaper facade', () => {
  it('edits a sheet-qualified input and returns formula readback restore proof', () => {
    const book = buildPricingBook()

    try {
      const proof = book.setCellAndReadback('Inputs!B2', 48, {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(true)
      expect(proof.editedCell).toBe('Inputs!B2')
      expect(proof.readbackRange).toBe('Summary!B2')
      expect(proof.beforeReadback.displayValues).toEqual([['48000']])
      expect(proof.afterReadback.displayValues).toEqual([['57600']])
      expect(proof.restoredReadback.displayValues).toEqual(proof.afterReadback.displayValues)
      expect(proof.persistedDocumentBytes).toBeGreaterThan(100)
      expect(proof.checks).toMatchObject({
        readbackChanged: true,
        restoredReadbackMatchesAfter: true,
        previousSerialized: 40,
        newSerialized: 48,
      })
    } finally {
      book.dispose()
    }
  })

  it('rejects writes outside the configured writable sheet list', () => {
    const book = buildPricingBook({ writableSheets: ['Inputs'] })

    try {
      expect(() =>
        book.setCellAndReadback('Summary!B2', '=Inputs!B2*Inputs!B3*2', {
          readbackRange: 'Summary!B2',
        }),
      ).toThrow('Sheet "Summary" is not writable')
    } finally {
      book.dispose()
    }
  })

  it('requires sheet-qualified addresses unless a default sheet is configured', () => {
    const strictBook = buildPricingBook()
    const defaultedBook = buildPricingBook({ defaultSheetName: 'Summary' })

    try {
      expect(() => strictBook.display('B2')).toThrow('Sheet-qualified WorkPaper cell address required: B2')
      expect(defaultedBook.display('B2')).toBe('48000')
    } finally {
      strictBook.dispose()
      defaultedBook.dispose()
    }
  })

  it('writes formulas through A1 addresses and preserves formula serialization', () => {
    const book = buildPricingBook()

    try {
      const proof = book.setCellAndReadback('Summary!B2', '=Inputs!B2*Inputs!B3*2', {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(true)
      expect(proof.after.formula).toBe('=Inputs!B2*Inputs!B3*2')
      expect(proof.after.serialized).toBe('=Inputs!B2*Inputs!B3*2')
      expect(proof.afterReadback.displayValues).toEqual([['96000']])
      expect(proof.restoredReadback.serialized).toEqual([['=Inputs!B2*Inputs!B3*2']])
    } finally {
      book.dispose()
    }
  })

  it('restores serialized A1 WorkPaper documents', () => {
    const source = buildPricingBook()

    try {
      source.set('Inputs!B2', 60)
      const restored = restoreA1WorkPaper(source.serialize())
      try {
        expect(restored.display('Summary!B2')).toBe('72000')
        expect(restored.readCell('Inputs!B2').serialized).toBe(60)
        expect(restored.serialize()).toContain('bilig.headless.work-paper.document.v1')
      } finally {
        restored.dispose()
      }
    } finally {
      source.dispose()
    }
  })

  it('provides concise multi-cell and JSON aliases for agent workflows', () => {
    const book = buildPricingBook()

    try {
      book.setMany({
        'Inputs!B2': 50,
        'Inputs!B3': 1500,
      })

      expect(book.display('Summary!B2')).toBe('75000')
      expect(Object.keys(book.readMany(['Inputs!B2', 'Summary!B2']))).toEqual(['Inputs!B2', 'Summary!B2'])

      const proof = book.editAndReadback('Inputs!B2', 60, {
        readbackRange: 'Summary!B2',
      })
      expect(proof.verified).toBe(true)
      expect(proof.afterReadback.displayValues).toEqual([['90000']])

      const restored = book.restoreJson(book.saveJson())
      try {
        expect(restored.display('Summary!B2')).toBe('90000')
      } finally {
        restored.dispose()
      }
    } finally {
      book.dispose()
    }
  })
})

function buildPricingBook(options: Parameters<typeof buildA1WorkPaper>[2] = {}) {
  return buildA1WorkPaper(
    {
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    },
    undefined,
    options,
  )
}
