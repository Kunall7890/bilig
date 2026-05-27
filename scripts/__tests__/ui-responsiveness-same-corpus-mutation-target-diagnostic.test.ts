import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import {
  SameCorpusCommittedStateMismatchError,
  type SameCorpusMutationTargetCommittedStatePhaseProof,
} from '../ui-responsiveness-same-corpus-committed-state-proof.ts'
import {
  SameCorpusMutationTargetCaptureDiagnosticError,
  sameCorpusMutationTargetFailureDiagnosticArtifactPath,
  writeSameCorpusMutationTargetCaptureFailureDiagnostic,
} from '../ui-responsiveness-same-corpus-mutation-target-diagnostic.ts'
import type { SameCorpusMutationTargetSelection } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import type { SameCorpusMutationTargetReadback, SameCorpusMutationTargetScreenshotProof } from '../ui-responsiveness-same-corpus-proof.ts'

describe('same-corpus mutation target failure diagnostics', () => {
  it('writes target-specific diagnostics for failed Google Sheets committed-state proof', () => {
    const outputPath = join(mkdtempSync(join(tmpdir(), 'bilig-mutation-diagnostic-')), 'capture.json')
    const target = targetSelection()
    const before = readback({ value: 'segment-5' })
    const lastReadback = readback({ fillColor: null, source: 'google-sheets-xlsx-export', value: 'segment-5' })
    const expectedReadback = readback({ fillColor: '#c9daf8', value: 'segment-5' })
    const afterCommittedStateProof = committedPhaseProof({
      phase: 'after',
      readback: lastReadback,
      target,
    })
    const error = new SameCorpusCommittedStateMismatchError({
      product: 'google-sheets',
      phase: 'after',
      sampleIndex: 0,
      workload: 'fill-format-change',
      sheetName: target.sheetName,
      sheetId: target.sheetId,
      targetRange: target.targetRange,
      expectedReadback,
      lastReadback,
      exportUrl: 'https://docs.google.com/spreadsheets/d/test/export?format=xlsx',
      lastArtifactPath: afterCommittedStateProof.artifactPath ?? null,
      lastArtifactSha256: afterCommittedStateProof.artifactSha256 ?? null,
      lastWorkbookSha256: afterCommittedStateProof.workbookSha256,
    })

    const relativePath = writeSameCorpusMutationTargetCaptureFailureDiagnostic({
      after: readback({ fillColor: 'rgb(11, 87, 208)', value: 'segment-5' }),
      afterCommittedStateProof,
      afterScreenshot: screenshotProof('after', target, readback({ fillColor: 'rgb(11, 87, 208)', value: 'segment-5' })),
      before,
      beforeCommittedStateProof: committedPhaseProof({ phase: 'before', readback: before, target }),
      beforeScreenshot: screenshotProof('before', target, before),
      error,
      failurePhase: 'capture-after-committed-state',
      intendedPayload: {
        kind: 'fill-color',
        expectedFillColor: '#c9daf8',
        swatchLabel: 'light blue 2',
      },
      operationStartedAt: 100,
      outputPath,
      product: 'google-sheets',
      restoreError: null,
      restored: null,
      restoredCommittedStateProof: null,
      restoredScreenshot: null,
      restoreStatus: 'restored',
      reselectError: null,
      reselectStatus: 'restored',
      sampleIndex: 0,
      target,
      visibleAfter: readback({ fillColor: 'rgb(11, 87, 208)', source: 'visible-grid-cell', value: 'segment-5' }),
      visibleAfterSelectedRange: 'C5',
      visibleRestored: null,
      visibleRestoredSelectedRange: null,
      workload: 'fill-format-change',
    })

    const artifactPath = sameCorpusMutationTargetFailureDiagnosticArtifactPath({
      outputPath,
      product: 'google-sheets',
      sampleIndex: 0,
      workload: 'fill-format-change',
    })
    const diagnostic: unknown = JSON.parse(readFileSync(artifactPath, 'utf8'))

    expect(relativePath).toContain('mutation-target-diagnostics/google-sheets-sample-1-failure.json')
    expect(diagnostic).toMatchObject({
      failurePhase: 'capture-after-committed-state',
      intendedPayload: {
        expectedFillColor: '#c9daf8',
      },
      targetRange: 'C5',
      committedStateMismatch: {
        expectedReadback: {
          fillColor: '#c9daf8',
        },
        lastReadback: {
          fillColor: null,
          source: 'google-sheets-xlsx-export',
        },
      },
    })
  })

  it('surfaces the diagnostic artifact path in the capture failure error', () => {
    const cause = new Error('committed XLSX export never showed the intended fill')
    const error = new SameCorpusMutationTargetCaptureDiagnosticError({
      artifactPath: 'capture.json.proof/same-corpus-fill-format-change/mutation-target-diagnostics/google-sheets-sample-1-failure.json',
      cause,
      failurePhase: 'capture-after-committed-state',
      product: 'google-sheets',
      sampleIndex: 0,
      targetRange: 'C5',
      workload: 'fill-format-change',
    })

    expect(error).toMatchObject({
      name: 'SameCorpusMutationTargetCaptureDiagnosticError',
      diagnosticArtifactPath:
        'capture.json.proof/same-corpus-fill-format-change/mutation-target-diagnostics/google-sheets-sample-1-failure.json',
      failurePhase: 'capture-after-committed-state',
    })
    expect(error.message).toContain('Diagnostic artifact:')
    expect(error.message).toContain('google-sheets-sample-1-failure.json')
    expect(error.cause).toBe(cause)
  })
})

function targetSelection(): SameCorpusMutationTargetSelection {
  return {
    endAddress: 'C5',
    sheetId: '0',
    sheetName: 'WideGrid',
    startAddress: 'C5',
    targetRange: 'C5',
  }
}

function readback(values: Partial<SameCorpusMutationTargetReadback>): SameCorpusMutationTargetReadback {
  return {
    fillColor: null,
    formula: null,
    source: 'visible-grid-cell',
    value: null,
    visibleText: values.value ?? values.formula ?? null,
    ...values,
  }
}

function screenshotProof(
  phase: SameCorpusMutationTargetScreenshotProof['phase'],
  target: SameCorpusMutationTargetSelection,
  semanticReadback: SameCorpusMutationTargetReadback,
): SameCorpusMutationTargetScreenshotProof {
  return {
    phase,
    product: 'google-sheets',
    scope: 'target-cell',
    sampleIndex: 0,
    sheetId: target.sheetId,
    sheetName: target.sheetName,
    targetRange: target.targetRange,
    workload: 'fill-format-change',
    screenshotPath: `.cache/${phase}.png`,
    screenshotSha256: '0'.repeat(64),
    semanticReadback,
  }
}

function committedPhaseProof(args: {
  readonly phase: SameCorpusMutationTargetCommittedStatePhaseProof['phase']
  readonly readback: SameCorpusMutationTargetReadback
  readonly target: SameCorpusMutationTargetSelection
}): SameCorpusMutationTargetCommittedStatePhaseProof {
  return {
    product: 'google-sheets',
    phase: args.phase,
    sampleIndex: 0,
    workload: 'fill-format-change',
    sheetName: args.target.sheetName,
    sheetId: args.target.sheetId,
    targetRange: args.target.targetRange,
    exportUrl: 'https://docs.google.com/spreadsheets/d/test/export?format=xlsx',
    capturedAtMs: 120,
    artifactPath: `.cache/${args.phase}.json`,
    artifactSha256: '1'.repeat(64),
    workbookByteSize: 1234,
    workbookSha256: '2'.repeat(64),
    readback: args.readback,
  }
}
