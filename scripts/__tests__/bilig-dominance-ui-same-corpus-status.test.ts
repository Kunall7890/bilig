import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { readUiSameCorpusCaptureArtifactStatus } from '../bilig-dominance-ui-same-corpus-status.ts'
import {
  formatSameCorpusUiSpeedGap,
  sameCorpusUiSpeedGaps,
  type SameCorpusUiSpeedGapCase,
  type SameCorpusUiSpeedMetric,
} from '../ui-responsiveness-same-corpus-speed-gaps.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'
import type { UiResponsivenessSameCorpusWorkload } from '../ui-responsiveness-same-corpus-workloads.ts'

describe('same-corpus UI dominance status', () => {
  it('surfaces stale local same-corpus capture artifacts before scorecard promotion', () => {
    const dir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-status-'))
    try {
      const capturePath = join(dir, 'same-corpus-capture.json')
      writeFileSync(
        capturePath,
        `${JSON.stringify({
          schemaVersion: 1,
          suite: 'ui-responsiveness-same-corpus-capture',
          sampleCount: 3,
          limitations: [],
          cases: [],
        })}\n`,
      )

      const status = readUiSameCorpusCaptureArtifactStatus({
        path: capturePath,
        displayPath: '.cache/ui-responsiveness/same-corpus-capture.json',
      })

      expect(status).toMatchObject({
        path: '.cache/ui-responsiveness/same-corpus-capture.json',
        exists: true,
        parseable: false,
        currentRunManifest: false,
        readyForScorecardGeneration: false,
        sampleCount: null,
        caseCount: null,
        tenXMeanAndP95CaseCount: null,
        missingRequiredWorkloads: [...requiredUiResponsivenessSameCorpusWorkloads],
        readinessErrors: [
          'legacy same-corpus capture artifact is missing the current runManifest contract',
          'Expected runManifest to be an object',
        ],
        runManifestInvalidReasons: [],
        legacyCapture: {
          schemaVersion: 1,
          suite: 'ui-responsiveness-same-corpus-capture',
          sampleCount: 3,
          caseCount: 0,
          capturedWorkloads: [],
          missingRequiredWorkloads: [...requiredUiResponsivenessSameCorpusWorkloads],
          pixelGridProofCaseCount: 0,
          tenXMeanAndP95CaseCount: 0,
          googleSheetsTenXRequirementSatisfied: false,
          contractGap: 'missing-run-manifest',
        },
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('does not infer legacy 10x status from ambiguous ratio fields', () => {
    const dir = mkdtempSync(join(tmpdir(), 'bilig-same-corpus-status-'))
    try {
      const capturePath = join(dir, 'same-corpus-capture.json')
      writeFileSync(
        capturePath,
        `${JSON.stringify({
          schemaVersion: 1,
          suite: 'ui-responsiveness-same-corpus-capture',
          sampleCount: 3,
          limitations: [],
          cases: requiredUiResponsivenessSameCorpusWorkloads.map((workload) => ({
            workload,
            meanRatio: 100,
            p95Ratio: 100,
            pixelGridProof: { captured: true },
          })),
        })}\n`,
      )

      const status = readUiSameCorpusCaptureArtifactStatus({
        path: capturePath,
        displayPath: '.cache/ui-responsiveness/same-corpus-capture.json',
      })

      expect(status.legacyCapture).toMatchObject({
        caseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        missingRequiredWorkloads: [],
        pixelGridProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        tenXMeanAndP95CaseCount: 0,
        googleSheetsTenXRequirementSatisfied: false,
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('ranks same-corpus UI speed gaps using the metric enforced by the 10x gate', () => {
    const proof = {
      cases: [
        sameCorpusSpeedGapCase({
          workload: 'fill-format-change',
          metric: 'operationResponseMs',
          biligMeanMs: 313.67,
          biligP95Ms: 328.94,
          googleMeanMs: 82.02,
          googleP95Ms: 87.47,
        }),
        sameCorpusSpeedGapCase({
          workload: 'scroll-horizontal',
          metric: 'scrollEventResponseMs',
          biligMeanMs: 999,
          biligP95Ms: 999,
          googleMeanMs: 1,
          googleP95Ms: 1,
          biligScrollMeanMs: 4,
          biligScrollP95Ms: 8,
          googleScrollMeanMs: 80,
          googleScrollP95Ms: 80,
        }),
      ],
    }

    const gaps = sameCorpusUiSpeedGaps(proof)

    expect(gaps).toHaveLength(1)
    expect(gaps[0]).toMatchObject({
      workload: 'fill-format-change',
      metric: 'operationResponseMs',
      meanRatio: expect.closeTo(0.2614850001594032, 8),
      p95Ratio: expect.closeTo(0.2659147564905454, 8),
      limitingAdditionalSpeedupTo10x: expect.closeTo(38.24311143623507, 8),
    })
    expect(formatSameCorpusUiSpeedGap(gaps[0])).toContain(
      'fill-format-change (operationResponseMs): current Google/Bilig mean 0.26x: p95 0.27x',
    )
  })
})

function sameCorpusSpeedGapCase(args: {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly metric: SameCorpusUiSpeedMetric
  readonly biligMeanMs: number
  readonly biligP95Ms: number
  readonly googleMeanMs: number
  readonly googleP95Ms: number
  readonly biligScrollMeanMs?: number
  readonly biligScrollP95Ms?: number
  readonly googleScrollMeanMs?: number
  readonly googleScrollP95Ms?: number
}): SameCorpusUiSpeedGapCase {
  return {
    workload: args.workload,
    tenXMeanAndP95Metric: args.metric,
    bilig: {
      operationResponseMs: {
        mean: args.biligMeanMs,
        p95: args.biligP95Ms,
      },
      ...(args.biligScrollMeanMs !== undefined && args.biligScrollP95Ms !== undefined
        ? {
            scrollEventResponseMs: {
              mean: args.biligScrollMeanMs,
              p95: args.biligScrollP95Ms,
            },
          }
        : {}),
    },
    googleSheets: {
      operationResponseMs: {
        mean: args.googleMeanMs,
        p95: args.googleP95Ms,
      },
      ...(args.googleScrollMeanMs !== undefined && args.googleScrollP95Ms !== undefined
        ? {
            scrollEventResponseMs: {
              mean: args.googleScrollMeanMs,
              p95: args.googleScrollP95Ms,
            },
          }
        : {}),
    },
  }
}
