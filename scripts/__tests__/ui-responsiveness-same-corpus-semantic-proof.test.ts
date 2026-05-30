import { describe, expect, it } from 'vitest'

import {
  validateSameCorpusProductSemanticUiProof,
  type SameCorpusMutationTargetProof,
  type SameCorpusProductSemanticUiProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import { sameCorpusMutationTargetRangeForSample } from '../ui-responsiveness-same-corpus-mutation-proof-page.ts'
import { sameCorpusMutationTargetProofSignature } from '../ui-responsiveness-same-corpus-mutation-target-signature.ts'
import type { UiResponsivenessSameCorpusProduct } from '../ui-responsiveness-same-corpus-scorecard-proof.ts'
import type { UiResponsivenessSameCorpusMutatingWorkload } from '../ui-responsiveness-same-corpus-workloads.ts'
import { sameCorpusEditVisibleCellValue, sameCorpusFormulaEditFormula } from '../ui-responsiveness-same-corpus-workload-runner.ts'

describe('same-corpus semantic UI mutation proof validation', () => {
  it('accepts mutation target proof tied to exact workload and sample screenshots', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(validSemanticProof(), {
      workload: 'edit-visible-cell',
      sampleCount: 3,
    })

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
  })

  it('rejects mutation target proof without a per-sample target proof signature', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) => {
          if (proof.sampleIndex !== 0) {
            return proof
          }
          const { targetProofSignature: _targetProofSignature, ...unsignedProof } = proof
          return unsignedProof as SameCorpusMutationTargetProof
        }),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell is missing target proof signature']),
    })
  })

  it('rejects mutation target proof when signed sample fields drift', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleAfter: Object.assign({}, proof.visibleAfter, {
                  value: 'stale-editor-text',
                  visibleText: 'stale-editor-text',
                }),
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell target proof signature does not match sample fields',
      ]),
    })
  })

  it('rejects mutation target proof without committed target proof timing', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { committedTargetProofMs: Number.NaN }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell is missing committed target proof timing',
      ]),
    })
  })

  it('rejects mutation target proof without explicit operation timing bounds', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { operationStartedAtMs: Number.NaN }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell is missing operation proof timing bounds',
      ]),
    })
  })

  it('rejects mutation target proof with non-monotonic operation timing bounds', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { postMutationProofCapturedAtMs: proof.operationStartedAtMs - 1 }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell has non-monotonic operation proof timing bounds',
      ]),
    })
  })

  it('rejects mutation target proof whose committed timing drifts from the proof window', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { committedTargetProofMs: proof.committedTargetProofMs + 20 }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed target timing does not match proof window',
      ]),
    })
  })

  it('rejects mutation target proof without before, after, and restored target-cell screenshots', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { targetScreenshots: null }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell is missing target-cell screenshots',
      ]),
    })
  })

  it('rejects mutation target screenshots that do not visibly change after mutation', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? staleAfterMutationTargetScreenshotProof(proof) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell before and after target screenshots are identical',
      ]),
    })
  })

  it('rejects mutation target screenshots that do not visibly restore after undo', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? staleRestoredMutationTargetScreenshotProof(proof) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell after and restored target screenshots are identical',
      ]),
    })
  })

  it('rejects mutation target screenshots whose explicit identity drifts from the sample', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.targetScreenshots
            ? Object.assign({}, proof, {
                targetScreenshots: {
                  ...proof.targetScreenshots,
                  after: {
                    ...proof.targetScreenshots.after,
                    product: 'google-sheets',
                    sampleIndex: 2,
                    sheetId: 'other-sheet-id',
                    sheetName: 'OtherSheet',
                    workload: 'formula-edit',
                  },
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell has mismatched after screenshot product',
        'semantic UI mutation target proof for edit-visible-cell has mismatched after screenshot workload',
        'semantic UI mutation target proof for edit-visible-cell has mismatched after screenshot sample',
        'semantic UI mutation target proof for edit-visible-cell has mismatched after screenshot sheet identity',
      ]),
    })
  })

  it('rejects mutation target proof whose top-level product drifts from the semantic proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { product: 'google-sheets' as const }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining(['semantic UI mutation target proof for edit-visible-cell has mismatched product']),
    })
  })

  it('rejects rendered selection text that merely contains the target range', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        selectedRange: 'visible selection C5 after editor commit',
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell target range does not match the rendered selection',
      ]),
    })
  })

  it('rejects product-branded mutation payloads instead of product-neutral edits', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                intendedPayload: { kind: 'cell-value' as const, value: 'bilig-same-corpus-1' },
                after: Object.assign({}, proof.after, { value: 'bilig-same-corpus-1', visibleText: 'bilig-same-corpus-1' }),
                visibleAfter: Object.assign({}, proof.visibleAfter, { value: 'bilig-same-corpus-1', visibleText: 'bilig-same-corpus-1' }),
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell uses a non-neutral intended value payload',
      ]),
    })
  })

  it('rejects mutation target proof whose post-mutation visible selection drifted away from the target', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { visibleAfterSelectedRange: 'B2' }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell post-mutation visible selected range does not match target range',
      ]),
    })
  })

  it('rejects mutation target proof whose restored visible selection drifted away from the target', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { visibleRestoredSelectedRange: 'B2' }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell restored visible selected range does not match target range',
      ]),
    })
  })

  it('rejects semantic UI proof without a concrete sheet id', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        sheetId: null,
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining(['semantic UI proof is missing sheet id']),
    })
  })

  it('rejects mutation target proof whose sheet id does not match the semantic proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { sheetId: 'stale-sheet-id' }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell target sheet id does not match semantic UI proof',
      ]),
    })
  })

  it('rejects mutation screenshots that are not tied to the workload and sample', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                screenshotPath: 'tmp/mutation-target/bilig-sample-1.png',
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the workload',
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the mutation target sample',
      ]),
    })
  })

  it('rejects reused mutation target screenshots across samples', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 1
            ? Object.assign({}, proof, {
                screenshotPath: 'tmp/same-corpus-wide-mixed-250k-edit-visible-cell/mutation-target/bilig-sample-1-after.png',
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell duplicates screenshot artifact path',
        'semantic UI mutation target proof for edit-visible-cell screenshot artifact path is not tied to the mutation target sample',
      ]),
    })
  })

  it('rejects Bilig mutation proof whose revisions are not bound to the target readback and scene proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                after: Object.assign({}, proof.after, {
                  capturedRevision: 'stale-readback-revision',
                  visibleSceneProofSha256: null,
                }),
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell authoritative revision does not match target readback',
        'semantic UI mutation target proof for edit-visible-cell is missing Bilig visible scene proof',
      ]),
    })
  })

  it('rejects Bilig mutation proof whose visible render revision does not match the target scene proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleRenderRevision: `bilig-visible-scene-sha256:${'f'.repeat(64)}`,
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell visible render revision does not match target scene proof',
      ]),
    })
  })

  it.each(['edit-visible-cell', 'formula-edit', 'fill-format-change'] as const)(
    'requires Bilig %s visible proof to come from rendered grid cell pixels',
    (workload) => {
      const accepted = validateSameCorpusProductSemanticUiProof(
        validSemanticProof({
          selectedRange: sameCorpusMutationTargetRangeForSample(workload, 0),
          mutationTargetProofs: [0, 1, 2].map((sampleIndex) => mutationTargetProof('bilig', workload, sampleIndex)),
        }),
        {
          workload,
          sampleCount: 3,
        },
      )
      const rejected = validateSameCorpusProductSemanticUiProof(
        validSemanticProof({
          selectedRange: sameCorpusMutationTargetRangeForSample(workload, 0),
          mutationTargetProofs: [0, 1, 2].map((sampleIndex) =>
            biligMutationTargetProofWithVisibleSource(workload, sampleIndex, 'visible-formula-bar'),
          ),
        }),
        {
          workload,
          sampleCount: 3,
        },
      )

      expect(accepted).toMatchObject({
        acceptedForCurrentScorecard: true,
        invalidReasons: [],
      })
      expect(rejected).toMatchObject({
        acceptedForCurrentScorecard: false,
        invalidReasons: expect.arrayContaining([
          `semantic UI mutation target proof for ${workload} visible render readback did not come from an accepted browser-visible source`,
        ]),
      })
    },
  )

  it('rejects Bilig formula proof when the rendered grid cell text is stale', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('formula-edit', 0),
        mutationTargetProofs: [0, 1, 2].map((sampleIndex) =>
          sampleIndex === 0
            ? Object.assign({}, mutationTargetProof('bilig', 'formula-edit', sampleIndex), {
                visibleAfter: {
                  value: 'stale result',
                  formula: null,
                  fillColor: null,
                  visibleText: 'stale result',
                  source: 'visible-grid-cell' as const,
                },
              })
            : mutationTargetProof('bilig', 'formula-edit', sampleIndex),
        ),
      }),
      {
        workload: 'formula-edit',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for formula-edit did not prove the rendered formula result',
      ]),
    })
  })

  it('rejects Bilig visible readback when native text is tied to a stale visible scene', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validSemanticProof({
        mutationTargetProofs: validMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? Object.assign({}, proof, {
                visibleAfter: {
                  ...proof.visibleAfter,
                  visibleSceneProofSha256: 'f'.repeat(64),
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell post-mutation visible render readback visible scene proof does not match target readback',
      ]),
    })
  })

  it('accepts Google Sheets mutation proof only when an independent XLSX export proves committed state', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(validGoogleSheetsSemanticProof(), {
      workload: 'edit-visible-cell',
      sampleCount: 3,
    })

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
  })

  it('rejects Google Sheets formula proof when XLSX export lacks the rendered formula result', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('formula-edit', 0),
        mutationTargetProofs: validGoogleSheetsFormulaMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.committedStateProof
            ? Object.assign({}, proof, {
                committedStateProof: {
                  ...proof.committedStateProof,
                  after: googleCommittedStatePhaseProof(
                    proof,
                    'after',
                    Object.assign({}, proof.after, { value: null, visibleText: proof.after.formula }),
                  ),
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'formula-edit',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for formula-edit committed-state after readback did not prove formula result',
      ]),
    })
  })

  it('rejects Google Sheets formula proof when target-grid proof is replaced by formula-bar chrome', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('formula-edit', 0),
        mutationTargetProofs: validGoogleSheetsFormulaMutationTargetProofs().map((proof) =>
          signedMutationTargetProof({
            ...proof,
            after: {
              ...proof.after,
              value: null,
              visibleText: proof.after.formula,
              source: 'visible-formula-bar',
            },
            visibleAfter: {
              ...proof.visibleAfter,
              value: null,
              formula: proof.after.formula,
              visibleText: proof.after.formula,
              source: 'visible-formula-bar',
            },
            visibleRestored: {
              ...proof.visibleRestored,
              value: proof.restored.value,
              formula: proof.restored.formula,
              visibleText: proof.restored.visibleText,
              source: 'visible-formula-bar',
            },
            targetScreenshots: proof.targetScreenshots
              ? {
                  ...proof.targetScreenshots,
                  after: {
                    ...proof.targetScreenshots.after,
                    semanticReadback: {
                      ...proof.targetScreenshots.after.semanticReadback,
                      value: null,
                      formula: proof.after.formula,
                      visibleText: proof.after.formula,
                      source: 'visible-formula-bar',
                    },
                  },
                  restored: {
                    ...proof.targetScreenshots.restored,
                    semanticReadback: {
                      ...proof.targetScreenshots.restored.semanticReadback,
                      value: proof.restored.value,
                      formula: proof.restored.formula,
                      visibleText: proof.restored.visibleText,
                      source: 'visible-formula-bar',
                    },
                  },
                }
              : proof.targetScreenshots,
          }),
        ),
      }),
      {
        workload: 'formula-edit',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for formula-edit target readback did not come from an accepted browser-visible source',
        'semantic UI mutation target proof for formula-edit visible render readback did not come from an accepted browser-visible source',
        'semantic UI mutation target proof for formula-edit after screenshot semantic readback did not come from an accepted browser-visible source',
      ]),
    })
  })

  it('rejects Google Sheets mutation proof without independent committed-state export proof', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? Object.assign({}, proof, { committedStateProof: null }) : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell is missing independent Google Sheets committed-state proof',
      ]),
    })
  })

  it('rejects Google Sheets committed-state proof without archived JSON artifacts', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.committedStateProof
            ? Object.assign({}, proof, {
                committedStateProof: {
                  ...proof.committedStateProof,
                  after: Object.assign({}, proof.committedStateProof.after, {
                    artifactPath: null,
                    artifactSha256: null,
                  }),
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed-state after proof is missing archive JSON artifact',
      ]),
    })
  })

  it('rejects Google Sheets editor-only text proof when exported committed target value is stale', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.committedStateProof
            ? Object.assign({}, proof, {
                committedStateProof: {
                  ...proof.committedStateProof,
                  after: googleCommittedStatePhaseProof(proof, 'after', proof.before),
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed-state export did not prove a before/after change',
        'semantic UI mutation target proof for edit-visible-cell committed-state after readback does not match target proof',
      ]),
    })
  })

  it('rejects Google Sheets committed-state proof when the target grid never visibly proves the mutation', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? signedMutationTargetProof({
                ...proof,
                visibleAfter: {
                  fillColor: null,
                  formula: null,
                  source: 'unknown',
                  value: null,
                  visibleText: null,
                },
                visibleRestored: {
                  fillColor: null,
                  formula: null,
                  source: 'unknown',
                  value: null,
                  visibleText: null,
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell has committed-state proof but is missing browser-visible target grid proof',
      ]),
    })
  })

  it('accepts Google Sheets edit proof when XLSX export owns semantics and target screenshots own canvas-grid visibility', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) => googleSheetsCanvasScreenshotBackedProof(proof)),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
  })

  it('accepts Google Sheets formula proof when XLSX export proves formula/result and target screenshots own canvas-grid visibility', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('formula-edit', 0),
        mutationTargetProofs: validGoogleSheetsFormulaMutationTargetProofs().map((proof) => googleSheetsCanvasScreenshotBackedProof(proof)),
      }),
      {
        workload: 'formula-edit',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: true,
      invalidReasons: [],
    })
  })

  it('rejects Google Sheets mutation proof when selected-target visible readback comes from the formula bar', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0
            ? signedMutationTargetProof({
                ...proof,
                visibleAfter: { ...proof.visibleAfter, source: 'visible-formula-bar' },
                visibleRestored: { ...proof.visibleRestored, source: 'visible-formula-bar' },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell visible render readback did not come from an accepted browser-visible source',
      ]),
    })
  })

  it('rejects Google Sheets target screenshots backed by selected-target formula-bar semantic readback', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.targetScreenshots
            ? signedMutationTargetProof({
                ...proof,
                targetScreenshots: {
                  ...proof.targetScreenshots,
                  after: {
                    ...proof.targetScreenshots.after,
                    semanticReadback: { ...proof.targetScreenshots.after.semanticReadback, source: 'visible-formula-bar' },
                  },
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell after screenshot semantic readback did not come from an accepted browser-visible source',
      ]),
    })
  })

  it('rejects Google Sheets committed-state exports captured outside the mutation proof windows', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) => {
          if (proof.sampleIndex !== 0 || !proof.committedStateProof) {
            return proof
          }
          return Object.assign({}, proof, {
            committedStateProof: {
              ...proof.committedStateProof,
              after: Object.assign({}, proof.committedStateProof.after, {
                capturedAtMs: proof.operationStartedAtMs - 1,
              }),
              before: Object.assign({}, proof.committedStateProof.before, {
                capturedAtMs: proof.operationStartedAtMs + 1,
              }),
              restored: Object.assign({}, proof.committedStateProof.restored, {
                capturedAtMs: proof.postMutationProofCapturedAtMs - 1,
              }),
            },
          })
        }),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed-state before export was not captured before mutation started',
        'semantic UI mutation target proof for edit-visible-cell committed-state after export was not captured inside the post-mutation proof window',
        'semantic UI mutation target proof for edit-visible-cell committed-state restored export was not captured inside the restore proof window',
      ]),
    })
  })

  it('rejects Google Sheets committed-state exports from a different spreadsheet', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) => {
          if (proof.sampleIndex !== 0 || !proof.committedStateProof) {
            return proof
          }
          return Object.assign({}, proof, {
            committedStateProof: {
              ...proof.committedStateProof,
              after: Object.assign({}, proof.committedStateProof.after, {
                exportUrl: 'https://docs.google.com/spreadsheets/d/other-spreadsheet/export?format=xlsx',
              }),
            },
          })
        }),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed-state after proof is from a different spreadsheet export URL',
      ]),
    })
  })

  it('rejects Google Sheets committed-state readback carrying browser-only proof metadata', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        mutationTargetProofs: validGoogleSheetsMutationTargetProofs().map((proof) => {
          if (proof.sampleIndex !== 0 || !proof.committedStateProof) {
            return proof
          }
          return Object.assign({}, proof, {
            committedStateProof: {
              ...proof.committedStateProof,
              after: Object.assign({}, proof.committedStateProof.after, {
                readback: {
                  ...proof.committedStateProof.after.readback,
                  capturedRevision: 'browser-authoritative-revision',
                  visibleSceneProofSha256: 'f'.repeat(64),
                },
              }),
            },
          })
        }),
      }),
      {
        workload: 'edit-visible-cell',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for edit-visible-cell committed-state after readback carries browser-only proof metadata',
      ]),
    })
  })

  it('rejects Google Sheets fill proof when toolbar fill changes but exported target fill does not', () => {
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('fill-format-change', 0),
        mutationTargetProofs: validGoogleSheetsFillMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 && proof.committedStateProof
            ? Object.assign({}, proof, {
                committedStateProof: {
                  ...proof.committedStateProof,
                  after: googleCommittedStatePhaseProof(proof, 'after', Object.assign({}, proof.after, { fillColor: null })),
                },
              })
            : proof,
        ),
      }),
      {
        workload: 'fill-format-change',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for fill-format-change committed-state export did not prove a before/after change',
        'semantic UI mutation target proof for fill-format-change committed-state after readback does not match target proof',
        'semantic UI mutation target proof for fill-format-change committed-state after export does not contain intended fill color',
      ]),
    })
  })

  it('rejects Google Sheets fill proof when the visible target only exposes selection-border color', () => {
    const selectionBorderColor = 'rgb(11, 87, 208)'
    const verdict = validateSameCorpusProductSemanticUiProof(
      validGoogleSheetsSemanticProof({
        selectedRange: sameCorpusMutationTargetRangeForSample('fill-format-change', 0),
        mutationTargetProofs: validGoogleSheetsFillMutationTargetProofs().map((proof) =>
          proof.sampleIndex === 0 ? fillProofWithVisibleAfterFill(proof, selectionBorderColor) : proof,
        ),
      }),
      {
        workload: 'fill-format-change',
        sampleCount: 3,
      },
    )

    expect(verdict).toMatchObject({
      acceptedForCurrentScorecard: false,
      invalidReasons: expect.arrayContaining([
        'semantic UI mutation target proof for fill-format-change rendered target cell does not show intended fill color',
        'semantic UI mutation target proof for fill-format-change rendered target fill does not match target readback',
      ]),
    })
  })
})

function validSemanticProof(overrides: Partial<SameCorpusProductSemanticUiProof> = {}): SameCorpusProductSemanticUiProof {
  return {
    product: 'bilig',
    captured: true,
    method: 'bilig-visible-semantic-readback',
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    selectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
    authoritativeRenderRevision: 'rev-3',
    visibleRenderRevision: 'scene-7',
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: validMutationTargetProofs(),
    evidence: [
      'sheetName=WideGrid',
      `selectedRange=${sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0)}`,
      'checkedCellCount=3',
      `screenshotSha256=${'a'.repeat(64)}`,
      'authoritativeRenderRevision=rev-3',
      'visibleRenderRevision=scene-7',
    ],
    ...overrides,
  }
}

function validMutationTargetProofs(): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => mutationTargetProof('bilig', 'edit-visible-cell', sampleIndex))
}

function validGoogleSheetsSemanticProof(overrides: Partial<SameCorpusProductSemanticUiProof> = {}): SameCorpusProductSemanticUiProof {
  return validSemanticProof({
    authoritativeRenderRevision: null,
    method: 'google-sheets-visible-semantic-readback',
    mutationTargetProofs: validGoogleSheetsMutationTargetProofs(),
    product: 'google-sheets',
    selectedRange: sameCorpusMutationTargetRangeForSample('edit-visible-cell', 0),
    sheetId: 'gid:12345',
    visibleRenderRevision: null,
    ...overrides,
  })
}

function validGoogleSheetsMutationTargetProofs(): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => googleSheetsMutationTargetProof(sampleIndex))
}

function validGoogleSheetsFormulaMutationTargetProofs(): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => {
    const proof = mutationTargetProof('google-sheets', 'formula-edit', sampleIndex)
    return signedMutationTargetProof({ ...proof, committedStateProof: googleCommittedStateProof(proof) })
  })
}

function validGoogleSheetsFillMutationTargetProofs(): SameCorpusMutationTargetProof[] {
  return [0, 1, 2].map((sampleIndex) => googleSheetsFillMutationTargetProof(sampleIndex))
}

function googleSheetsMutationTargetProof(sampleIndex: number): SameCorpusMutationTargetProof {
  const proof = mutationTargetProof('google-sheets', 'edit-visible-cell', sampleIndex)
  return signedMutationTargetProof({ ...proof, committedStateProof: googleCommittedStateProof(proof) })
}

function googleSheetsFillMutationTargetProof(sampleIndex: number): SameCorpusMutationTargetProof {
  const proof = {
    ...fillMutationTargetProof(sampleIndex, 'visible-grid-cell'),
    product: 'google-sheets' as const,
    sheetId: 'gid:12345',
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-fill-format-change/mutation-target/google-sheets-sample-${String(
      sampleIndex + 1,
    )}-after.png`,
    targetScreenshots: mutationTargetScreenshots('google-sheets', 'fill-format-change', sampleIndex),
  }
  return signedMutationTargetProof({ ...proof, committedStateProof: googleCommittedStateProof(proof) })
}

function googleSheetsCanvasScreenshotBackedProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (!proof.committedStateProof || !proof.targetScreenshots) {
    return proof
  }
  const screenshotReadback: SameCorpusMutationTargetProof['visibleAfter'] = {
    fillColor: null,
    formula: null,
    source: 'visible-grid-target-screenshot',
    value: null,
    visibleText: null,
  }
  return signedMutationTargetProof({
    ...proof,
    before: proof.committedStateProof.before.readback,
    after: proof.committedStateProof.after.readback,
    restored: proof.committedStateProof.restored.readback,
    visibleAfter: screenshotReadback,
    visibleRestored: screenshotReadback,
    targetScreenshots: {
      before: { ...proof.targetScreenshots.before, semanticReadback: screenshotReadback },
      after: { ...proof.targetScreenshots.after, semanticReadback: screenshotReadback },
      restored: { ...proof.targetScreenshots.restored, semanticReadback: screenshotReadback },
    },
  })
}

function mutationTargetProof(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof {
  const operationStartedAtMs = 1000 + sampleIndex * 100
  const committedTargetProofMs = 40 + sampleIndex
  const visibleTargetRenderMs = 8 + sampleIndex
  const committedStateValidationMs = committedTargetProofMs - visibleTargetRenderMs
  const restoreValidationMs = 80
  const authoritativeSource = product === 'bilig' ? 'bilig-authoritative-range' : 'visible-grid-cell'
  return signedMutationTargetProof({
    product,
    sampleIndex,
    committedTargetProofMs,
    visibleTargetRenderMs,
    committedStateValidationMs,
    restoreValidationMs,
    operationStartedAtMs,
    visibleTargetRenderCapturedAtMs: operationStartedAtMs + visibleTargetRenderMs,
    postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs,
    restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs + restoreValidationMs,
    workload,
    intendedOperation: workload,
    intendedPayload: mutationTargetIntendedPayload(workload, sampleIndex),
    sheetName: 'WideGrid',
    sheetId: product === 'google-sheets' ? 'gid:12345' : 'sheet-wide-grid',
    targetRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
    before: {
      value: mutationTargetBeforeValue(workload),
      formula: null,
      fillColor: null,
      visibleText: mutationTargetBeforeValue(workload),
      source: authoritativeSource,
      capturedRevision: `before-readback-${String(sampleIndex + 1)}`,
      ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
    },
    after: Object.assign(mutationTargetAfterReadback(workload, sampleIndex), {
      source: authoritativeSource,
      capturedRevision: authoritativeReadbackRevision(sampleIndex),
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    }),
    restored: {
      value: mutationTargetBeforeValue(workload),
      formula: null,
      fillColor: null,
      visibleText: mutationTargetBeforeValue(workload),
      source: authoritativeSource,
      capturedRevision: `restored-readback-${String(sampleIndex + 1)}`,
      ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
    },
    visibleAfter: mutationTargetVisibleReadback(product, workload, 'after', sampleIndex),
    visibleAfterSelectedRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
    visibleRestored: mutationTargetVisibleReadback(product, workload, 'before', sampleIndex),
    visibleRestoredSelectedRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision: visibleRenderRevision(sampleIndex),
    targetScreenshots: mutationTargetScreenshots(product, workload, sampleIndex),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, 'after'),
    undoRestoreStatus: 'verified',
  })
}

function biligMutationTargetProofWithVisibleSource(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
  source: 'visible-formula-bar' | 'visible-grid-cell',
): SameCorpusMutationTargetProof {
  const proof = mutationTargetProof('bilig', workload, sampleIndex)
  return signedMutationTargetProof({
    ...proof,
    visibleAfter: { ...proof.visibleAfter, source },
    visibleRestored: { ...proof.visibleRestored, source },
  })
}

function mutationTargetIntendedPayload(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['intendedPayload'] {
  if (workload === 'formula-edit') {
    return { kind: 'formula', formula: sameCorpusFormulaEditFormula(sampleIndex) }
  }
  if (workload === 'fill-format-change') {
    const fillColor = fillColorForSample(sampleIndex)
    return { kind: 'fill-color', expectedFillColor: fillColor, swatchLabel: `swatch-${String(sampleIndex + 1)}` }
  }
  return { kind: 'cell-value', value: sameCorpusEditVisibleCellValue(sampleIndex) }
}

function mutationTargetBeforeValue(workload: UiResponsivenessSameCorpusMutatingWorkload): string {
  return workload === 'formula-edit' ? '1' : 'metric-1'
}

function mutationTargetAfterReadback(
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): Pick<SameCorpusMutationTargetProof['after'], 'fillColor' | 'formula' | 'value' | 'visibleText'> {
  if (workload === 'formula-edit') {
    const value = String(sampleIndex + 2)
    return {
      value,
      formula: sameCorpusFormulaEditFormula(sampleIndex),
      fillColor: null,
      visibleText: value,
    }
  }
  if (workload === 'fill-format-change') {
    return {
      value: 'metric-1',
      formula: null,
      fillColor: fillColorForSample(sampleIndex),
      visibleText: 'metric-1',
    }
  }
  return {
    value: sameCorpusEditVisibleCellValue(sampleIndex),
    formula: null,
    fillColor: null,
    visibleText: sameCorpusEditVisibleCellValue(sampleIndex),
  }
}

function mutationTargetVisibleReadback(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  phase: 'before' | 'after',
  sampleIndex: number,
): SameCorpusMutationTargetProof['visibleAfter'] {
  const source = 'visible-grid-cell'
  if (phase === 'before') {
    const value = mutationTargetBeforeValue(workload)
    return {
      value,
      formula: null,
      fillColor: null,
      visibleText: value,
      source,
      ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
    }
  }
  const after = mutationTargetAfterReadback(workload, sampleIndex)
  return {
    value: after.value,
    formula: null,
    fillColor: after.fillColor,
    visibleText: after.visibleText,
    source,
    ...(product === 'bilig' ? { visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex) } : {}),
  }
}

function fillMutationTargetProof(
  sampleIndex: number,
  visibleSource: 'visible-formula-bar' | 'visible-grid-cell',
): SameCorpusMutationTargetProof {
  const fillColor = fillColorForSample(sampleIndex)
  const operationStartedAtMs = 2000 + sampleIndex * 100
  const committedTargetProofMs = 60 + sampleIndex
  const visibleTargetRenderMs = 12 + sampleIndex
  const committedStateValidationMs = committedTargetProofMs - visibleTargetRenderMs
  const restoreValidationMs = 80
  const targetRange = sameCorpusMutationTargetRangeForSample('fill-format-change', sampleIndex)
  return signedMutationTargetProof({
    product: 'bilig',
    sampleIndex,
    committedTargetProofMs,
    visibleTargetRenderMs,
    committedStateValidationMs,
    restoreValidationMs,
    operationStartedAtMs,
    visibleTargetRenderCapturedAtMs: operationStartedAtMs + visibleTargetRenderMs,
    postMutationProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs,
    restoreProofCapturedAtMs: operationStartedAtMs + committedTargetProofMs + restoreValidationMs,
    workload: 'fill-format-change',
    intendedOperation: 'fill-format-change',
    intendedPayload: {
      kind: 'fill-color',
      expectedFillColor: fillColor,
      swatchLabel: `swatch-${String(sampleIndex + 1)}`,
    },
    sheetName: 'WideGrid',
    sheetId: 'sheet-wide-grid',
    targetRange,
    before: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `before-readback-${String(sampleIndex + 1)}`,
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    after: {
      value: 'metric-1',
      formula: null,
      fillColor,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: authoritativeReadbackRevision(sampleIndex),
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    restored: {
      value: 'metric-1',
      formula: null,
      fillColor: null,
      visibleText: 'metric-1',
      source: 'bilig-authoritative-range',
      capturedRevision: `restored-readback-${String(sampleIndex + 1)}`,
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    visibleAfter: {
      value: null,
      formula: null,
      fillColor,
      visibleText: null,
      source: visibleSource,
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    visibleAfterSelectedRange: targetRange,
    visibleRestored: {
      value: null,
      formula: null,
      fillColor: null,
      visibleText: null,
      source: visibleSource,
      visibleSceneProofSha256: visibleSceneProofSha256(sampleIndex),
    },
    visibleRestoredSelectedRange: targetRange,
    authoritativeReadbackRevision: authoritativeReadbackRevision(sampleIndex),
    visibleRenderRevision: visibleRenderRevision(sampleIndex),
    targetScreenshots: mutationTargetScreenshots('bilig', 'fill-format-change', sampleIndex),
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-fill-format-change/mutation-target/bilig-sample-${String(sampleIndex + 1)}-after.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, 'after'),
    undoRestoreStatus: 'verified',
  })
}

function signedMutationTargetProof(
  proof: Omit<SameCorpusMutationTargetProof, 'targetProofSignature'> | SameCorpusMutationTargetProof,
): SameCorpusMutationTargetProof {
  return {
    ...proof,
    targetProofSignature: sameCorpusMutationTargetProofSignature(proof),
  }
}

function mutationTargetScreenshots(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
): SameCorpusMutationTargetProof['targetScreenshots'] {
  return {
    before: mutationTargetScreenshot(product, workload, sampleIndex, 'before'),
    after: mutationTargetScreenshot(product, workload, sampleIndex, 'after'),
    restored: mutationTargetScreenshot(product, workload, sampleIndex, 'restored'),
  }
}

function mutationTargetScreenshot(
  product: UiResponsivenessSameCorpusProduct,
  workload: UiResponsivenessSameCorpusMutatingWorkload,
  sampleIndex: number,
  phase: 'before' | 'after' | 'restored',
): NonNullable<SameCorpusMutationTargetProof['targetScreenshots']>['before'] {
  return {
    phase,
    product,
    scope: 'target-cell',
    sampleIndex,
    sheetId: product === 'google-sheets' ? 'gid:12345' : 'sheet-wide-grid',
    sheetName: 'WideGrid',
    targetRange: sameCorpusMutationTargetRangeForSample(workload, sampleIndex),
    workload,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${String(sampleIndex + 1)}-${phase}.png`,
    screenshotSha256: mutationTargetScreenshotSha256(sampleIndex, phase),
    semanticReadback: mutationTargetVisibleReadback(product, workload, phase === 'after' ? 'after' : 'before', sampleIndex),
  }
}

function staleAfterMutationTargetScreenshotProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (!proof.targetScreenshots) {
    return proof
  }
  const beforeHash = proof.targetScreenshots.before.screenshotSha256
  return {
    ...proof,
    screenshotSha256: beforeHash,
    targetScreenshots: {
      ...proof.targetScreenshots,
      after: {
        ...proof.targetScreenshots.after,
        screenshotSha256: beforeHash,
      },
    },
  }
}

function staleRestoredMutationTargetScreenshotProof(proof: SameCorpusMutationTargetProof): SameCorpusMutationTargetProof {
  if (!proof.targetScreenshots) {
    return proof
  }
  return {
    ...proof,
    targetScreenshots: {
      ...proof.targetScreenshots,
      restored: {
        ...proof.targetScreenshots.restored,
        screenshotSha256: proof.targetScreenshots.after.screenshotSha256,
      },
    },
  }
}

function fillProofWithVisibleAfterFill(proof: SameCorpusMutationTargetProof, fillColor: string): SameCorpusMutationTargetProof {
  const visibleAfter = { ...proof.visibleAfter, fillColor }
  if (!proof.targetScreenshots) {
    return { ...proof, visibleAfter }
  }
  return {
    ...proof,
    visibleAfter,
    targetScreenshots: {
      ...proof.targetScreenshots,
      after: {
        ...proof.targetScreenshots.after,
        semanticReadback: visibleAfter,
      },
    },
  }
}

function mutationTargetScreenshotSha256(sampleIndex: number, phase: 'before' | 'after' | 'restored'): string {
  const hexChars = '0123456789abcdef'
  const phaseOffset = phase === 'before' ? 1 : phase === 'after' ? 5 : 9
  return hexChars[(sampleIndex + phaseOffset) % hexChars.length]?.repeat(64) ?? '0'.repeat(64)
}

function googleCommittedStateProof(
  sample: SameCorpusMutationTargetProof,
): NonNullable<SameCorpusMutationTargetProof['committedStateProof']> {
  return {
    after: googleCommittedStatePhaseProof(sample, 'after', sample.after),
    before: googleCommittedStatePhaseProof(sample, 'before', sample.before),
    product: 'google-sheets',
    restored: googleCommittedStatePhaseProof(sample, 'restored', sample.restored),
    sampleIndex: sample.sampleIndex,
    sheetId: sample.sheetId,
    sheetName: sample.sheetName,
    source: 'google-sheets-xlsx-export',
    targetRange: sample.targetRange,
    workload: sample.intendedOperation,
  }
}

function googleCommittedStatePhaseProof(
  sample: SameCorpusMutationTargetProof,
  phase: 'before' | 'after' | 'restored',
  readback: SameCorpusMutationTargetProof['before'],
): NonNullable<SameCorpusMutationTargetProof['committedStateProof']>['before'] {
  const capturedAtMs =
    phase === 'before'
      ? sample.operationStartedAtMs - 10
      : phase === 'after'
        ? sample.operationStartedAtMs + Math.max(1, sample.committedTargetProofMs / 2)
        : sample.postMutationProofCapturedAtMs + 10
  return {
    artifactPath: `tmp/same-corpus-wide-mixed-250k-${sample.intendedOperation}/committed-state/google-sheets-sample-${String(
      sample.sampleIndex + 1,
    )}-${phase}.json`,
    artifactSha256: mutationTargetScreenshotSha256(sample.sampleIndex, phase),
    capturedAtMs,
    exportUrl: 'https://docs.google.com/spreadsheets/d/test-spreadsheet/export?format=xlsx',
    phase,
    product: 'google-sheets',
    readback: googleCommittedStateReadback(readback),
    sampleIndex: sample.sampleIndex,
    sheetId: sample.sheetId,
    sheetName: sample.sheetName,
    targetRange: sample.targetRange,
    workbookByteSize: 123456 + sample.sampleIndex,
    workbookSha256: mutationTargetScreenshotSha256(sample.sampleIndex, phase),
    workload: sample.intendedOperation,
  }
}

function googleCommittedStateReadback(readback: SameCorpusMutationTargetProof['before']): SameCorpusMutationTargetProof['before'] {
  return {
    value: readback.value,
    formula: readback.formula,
    fillColor: readback.fillColor,
    visibleText: readback.visibleText,
    source: 'google-sheets-xlsx-export',
  }
}

function fillColorForSample(sampleIndex: number): string {
  const colors = ['#00ff00', '#ffff00', '#ff00ff'] as const
  return colors[sampleIndex % colors.length]
}

function authoritativeReadbackRevision(sampleIndex: number): string {
  return `authoritative-readback-${String(sampleIndex + 1)}`
}

function visibleRenderRevision(sampleIndex: number): string {
  return `bilig-visible-scene-sha256:${visibleSceneProofSha256(sampleIndex)}`
}

function visibleSceneProofSha256(sampleIndex: number): string {
  return String(sampleIndex + 1)
    .repeat(64)
    .slice(0, 64)
}
