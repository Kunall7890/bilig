import { describe, expect, it } from 'vitest'
import { Effect } from 'effect'
import {
  assertMergeRangesDoNotOverlap,
  canonicalWorkbookFilterRange,
  canonicalWorkbookFilterRangeOnSheet,
  canonicalWorkbookRangeOnSheet,
  metadataEffect,
  normalizeMetadataKey,
  renameDataValidationSourceSheet,
} from '../workbook-metadata-service-helpers.js'
import type { WorkbookDataValidationRecord } from '../workbook-metadata-types.js'

describe('workbook metadata service helpers', () => {
  it('renames only sheet-scoped data validation list sources', () => {
    const base: WorkbookDataValidationRecord = {
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' },
      rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Old', startAddress: 'B1', endAddress: 'B3' } },
    }

    expect(renameDataValidationSourceSheet(base, 'Old', 'New').rule).toEqual({
      kind: 'list',
      source: { kind: 'range-ref', sheetName: 'New', startAddress: 'B1', endAddress: 'B3' },
    })
    expect(
      renameDataValidationSourceSheet(
        { ...base, rule: { kind: 'list', source: { kind: 'cell-ref', sheetName: 'Other', address: 'C1' } } },
        'Old',
        'New',
      ).rule,
    ).toEqual({ kind: 'list', source: { kind: 'cell-ref', sheetName: 'Other', address: 'C1' } })
    expect(
      renameDataValidationSourceSheet({ ...base, rule: { kind: 'list', source: { kind: 'named-range', name: 'Choices' } } }, 'Old', 'New')
        .rule,
    ).toEqual({ kind: 'list', source: { kind: 'named-range', name: 'Choices' } })
    expect(renameDataValidationSourceSheet({ ...base, rule: { kind: 'any' } }, 'Old', 'New').rule).toEqual({ kind: 'any' })
  })

  it('canonicalizes ranges and rejects overlapping merges', () => {
    expect(canonicalWorkbookRangeOnSheet('Sheet1', { sheetName: 'Ignored', startAddress: 'C3', endAddress: 'A1' })).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
    })
    expect(
      canonicalWorkbookFilterRange({
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'A1',
        criteria: [{ colId: 1, filters: { values: ['Open'] } }],
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
      criteria: [{ colId: 1, filters: { values: ['Open'] } }],
    })
    expect(canonicalWorkbookFilterRangeOnSheet('Sheet2', { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'A1' })).toEqual({
      sheetName: 'Sheet2',
      startAddress: 'A1',
      endAddress: 'B2',
    })
    expect(() =>
      assertMergeRangesDoNotOverlap([
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      ]),
    ).toThrow('Merged ranges cannot overlap')
    expect(() =>
      assertMergeRangesDoNotOverlap([
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
      ]),
    ).not.toThrow()
  })

  it('normalizes metadata keys and preserves stable metadata effect errors', () => {
    expect(normalizeMetadataKey('  author  ')).toBe('author')
    expect(() => normalizeMetadataKey('   ')).toThrow('Workbook metadata keys must be non-empty')
    expect(Effect.runSync(metadataEffect('fallback', () => 42))).toBe(42)
    expect(() =>
      Effect.runSync(
        metadataEffect('fallback', () => {
          throw new Error('specific')
        }),
      ),
    ).toThrow('specific')
    expect(() =>
      Effect.runSync(
        metadataEffect('fallback', () => {
          throw 'raw'
        }),
      ),
    ).toThrow('fallback')
  })
})
