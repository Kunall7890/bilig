import { describe, expect, it } from 'vitest'

import type { SameCorpusMutationTargetCommittedStatePhaseProof } from '../ui-responsiveness-same-corpus-committed-state-proof.ts'
import type { SameCorpusMutationTargetSelection } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import {
  assertSameCorpusMutationTargetPreflightProof,
  sameCorpusMutationTargetPreflightInvalidReasons,
} from '../ui-responsiveness-same-corpus-mutation-preflight.ts'
import type {
  SameCorpusMutationTargetReadback,
  SameCorpusMutationTargetScreenshotProof,
} from '../ui-responsiveness-same-corpus-semantic-proof.ts'

const sha256 = 'a'.repeat(64)
const target = {
  endAddress: 'C5',
  sheetId: 'gid-1',
  sheetName: 'WideGrid',
  startAddress: 'C5',
  targetRange: 'C5',
} as const satisfies SameCorpusMutationTargetSelection
const beforeReadback = {
  fillColor: '#ffffff',
  formula: null,
  source: 'visible-grid-cell',
  value: 'before',
  visibleText: 'before',
} as const satisfies SameCorpusMutationTargetReadback
const biligBeforeReadback = {
  ...beforeReadback,
  source: 'bilig-authoritative-range',
} as const satisfies SameCorpusMutationTargetReadback

describe('same-corpus mutation target preflight proof', () => {
  it('accepts Bilig proof only when the target-cell screenshot owns the before state', () => {
    expect(() =>
      assertSameCorpusMutationTargetPreflightProof({
        before: biligBeforeReadback,
        beforeCommittedStateProof: null,
        beforeScreenshot: beforeScreenshotProof('bilig', beforeReadback),
        product: 'bilig',
        sampleIndex: 0,
        target,
        visibleBefore: beforeReadback,
        workload: 'edit-visible-cell',
      }),
    ).not.toThrow()
  })

  it('rejects a fallback screenshot before mutating the workbook', () => {
    const invalidReasons = sameCorpusMutationTargetPreflightInvalidReasons({
      before: beforeReadback,
      beforeCommittedStateProof: committedBeforeProof(beforeReadback),
      beforeScreenshot: { ...beforeScreenshotProof('google-sheets', beforeReadback), scope: 'visible-grid-fallback' },
      product: 'google-sheets',
      sampleIndex: 0,
      target,
      visibleBefore: beforeReadback,
      workload: 'edit-visible-cell',
    })

    expect(invalidReasons).toContain('before screenshot proof is not scoped to the target cell')
  })

  it('rejects Google Sheets mutation samples without independent committed-state proof', () => {
    expect(() =>
      assertSameCorpusMutationTargetPreflightProof({
        before: beforeReadback,
        beforeCommittedStateProof: null,
        beforeScreenshot: beforeScreenshotProof('google-sheets', beforeReadback),
        product: 'google-sheets',
        sampleIndex: 0,
        target,
        visibleBefore: beforeReadback,
        workload: 'edit-visible-cell',
      }),
    ).toThrow('missing independent Google Sheets before committed-state proof')
  })

  it('accepts Google Sheets proof when target screenshot and committed workbook readback agree', () => {
    expect(
      sameCorpusMutationTargetPreflightInvalidReasons({
        before: beforeReadback,
        beforeCommittedStateProof: committedBeforeProof({
          ...beforeReadback,
          source: 'google-sheets-xlsx-export',
        }),
        beforeScreenshot: beforeScreenshotProof('google-sheets', beforeReadback),
        product: 'google-sheets',
        sampleIndex: 0,
        target,
        visibleBefore: beforeReadback,
        workload: 'edit-visible-cell',
      }),
    ).toEqual([])
  })
})

function beforeScreenshotProof(
  product: 'bilig' | 'google-sheets',
  semanticReadback: SameCorpusMutationTargetReadback,
): SameCorpusMutationTargetScreenshotProof {
  return {
    phase: 'before',
    product,
    sampleIndex: 0,
    scope: 'target-cell',
    screenshotPath: `.cache/proof/${product}-before.png`,
    screenshotSha256: sha256,
    semanticReadback,
    sheetId: target.sheetId,
    sheetName: target.sheetName,
    targetRange: target.targetRange,
    workload: 'edit-visible-cell',
  }
}

function committedBeforeProof(readback: SameCorpusMutationTargetReadback): SameCorpusMutationTargetCommittedStatePhaseProof {
  return {
    artifactPath: '.cache/proof/google-before.json',
    artifactSha256: sha256,
    capturedAtMs: 1,
    exportUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/export?format=xlsx',
    phase: 'before',
    product: 'google-sheets',
    readback,
    sampleIndex: 0,
    sheetId: target.sheetId,
    sheetName: target.sheetName,
    targetRange: target.targetRange,
    workbookByteSize: 128,
    workbookSha256: sha256,
    workload: 'edit-visible-cell',
  }
}
