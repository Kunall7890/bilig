import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import { hasUiResponsivenessSameCorpusTenXGap } from '../bilig-dominance-completion-audit.ts'
import {
  buildSameCorpusCaptureRunManifest,
  buildSameCorpusProof,
  assertUiResponsivenessLiveBrowserRunAllowed,
  parseUiResponsivenessLiveBrowserCliArgs,
  parseUiResponsivenessLiveBrowserScorecard,
  validateUiResponsivenessLiveBrowserScorecard,
  type SameCorpusCapture,
  type UiResponsivenessLiveBrowserScorecard,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'
import { buildSameCorpusFingerprint } from '../ui-responsiveness-same-corpus-fingerprint.ts'
import {
  validateSameCorpusProductPixelGridProof,
  type SameCorpusPixelGridProof,
  type SameCorpusProductPixelGridProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadRequiresScrollEventEvidence,
  type UiResponsivenessSameCorpusWorkload,
} from '../ui-responsiveness-same-corpus-workloads.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const sameCorpusFixtureFingerprint = buildSameCorpusFingerprint(buildWorkbookBenchmarkCorpus('wide-mixed-250k')).corpusFingerprint
const googleSheetsSourceWorkbookSha256 = '1'.repeat(64)
const microsoftExcelWebSourceWorkbookSha256 = '2'.repeat(64)

describe('UI responsiveness live browser scorecard', () => {
  it('validates the checked-in browser timing artifact', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(scorecard.summary).toMatchObject({
      directBrowserTimingCaptured: true,
      allRequiredCasesPassed: true,
      requiredVendorCount: 2,
      capturedVendors: ['google-sheets', 'microsoft-excel-web'],
    })
    expect(scorecard.cases.map((entry) => entry.id)).toEqual(['google-sheets-public-grid-scroll', 'microsoft-excel-web-public-xlsx-scroll'])
    expect(scorecard.cases.every((entry) => entry.sampleCount >= 3 && entry.limitations.length > 0)).toBe(true)
    expect(scorecard.sameCorpusProof).toMatchObject({
      captured: false,
      evidenceKind: 'not-captured',
      requiredProductCount: 2,
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      coveredCorpusCaseIds: [],
    })
    expect(scorecard.sameCorpusProof.cases).toEqual([])
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBe(
      scorecard.sameCorpusProof.cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length,
    )
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBe(0)
    expect(scorecard.sameCorpusProof.runManifest).toMatchObject({
      contractVersion: 'same-corpus-ui-v3',
      caseCount: 0,
      strictRenderedGridProofCaseCount: 0,
      legacyInsufficientRenderedGridProofCaseCount: 0,
      tenXMeanAndP95CaseCount: 0,
      currentContractEvidenceComplete: false,
      googleSheetsTenXRequirementSatisfied: false,
    })
    expect(scorecard.sameCorpusProof.runManifest?.invalidReasons).toContain('strict rendered-grid proof covers 0/9 cases')
    expect(scorecard.sameCorpusProof.limitations).toContain(
      'Same-corpus live browser timing against Bilig and Google Sheets has not been captured yet.',
    )
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
  })

  it('parses live browser scorecard CLI options', () => {
    expect(parseUiResponsivenessLiveBrowserCliArgs(['--check', '--capture', 'tmp/same-corpus-capture.json'])).toEqual({
      isCheckMode: true,
      capturePath: 'tmp/same-corpus-capture.json',
    })
  })

  it('rejects blank live browser capture paths', () => {
    expect(() => parseUiResponsivenessLiveBrowserCliArgs(['--capture', '   '])).toThrow('Missing value after --capture')
  })

  it('rejects live browser capture paths that consume the next flag', () => {
    expect(() => parseUiResponsivenessLiveBrowserCliArgs(['--capture', '--check'])).toThrow('Missing value after --capture')
  })

  it('rejects missing incumbent vendors', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      summary: {
        ...scorecard.summary,
        capturedVendors: ['google-sheets'],
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness live browser scorecard is missing vendor: microsoft-excel-web',
    )
  })

  it('blocks live browser scorecard generation while the local resource guard is active', () => {
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-ui-browser-live-guard-`)
    const coordinationDir = resolve(rootDir, '.agent-coordination')
    mkdirSync(coordinationDir)
    writeFileSync(
      resolve(coordinationDir, '20260508T092619Z-codex-memory-pressure-stop.md'),
      '# Memory pressure stop\n\nStatus: active on 2026-05-08T09:26:19Z.\n',
    )

    expect(() => assertUiResponsivenessLiveBrowserRunAllowed(rootDir, {})).toThrow(
      /Refusing to start UI responsiveness live browser scorecard generation/u,
    )
    expect(() => assertUiResponsivenessLiveBrowserRunAllowed(rootDir, { BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD: '1' })).not.toThrow()
  })

  it('derives same-corpus 10x ratios from operation and scroll-event samples', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())

    expect(proof).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredProductCount: 2,
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      tenXMeanAndP95CaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      coveredCorpusCaseIds: ['wide-mixed-250k'],
      runManifest: {
        contractVersion: 'same-corpus-ui-v3',
        requiredProducts: ['bilig', 'google-sheets'],
        requiredWorkloads: requiredUiResponsivenessSameCorpusWorkloads,
        capturedWorkloads: requiredUiResponsivenessSameCorpusWorkloads,
        corpusCaseIds: ['wide-mixed-250k'],
        corpusFingerprints: [sameCorpusFixtureFingerprint],
        productSourceWorkbookFingerprints: [
          {
            product: 'bilig',
            method: 'bilig-benchmark-state',
            source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
            sourceWorkbookSha256: sameCorpusFixtureFingerprint.snapshotSha256,
          },
          {
            product: 'google-sheets',
            method: 'google-sheets-xlsx-export',
            source: 'https://docs.google.com/spreadsheets/d/example',
            sourceWorkbookSha256: googleSheetsSourceWorkbookSha256,
          },
          {
            product: 'microsoft-excel-web',
            method: 'microsoft-excel-web-source-xlsx',
            source: 'https://view.officeapps.live.com/op/view.aspx?src=example',
            sourceWorkbookSha256: microsoftExcelWebSourceWorkbookSha256,
          },
        ],
        materializedCellCounts: [250000],
        sampleCount: 3,
        caseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        strictRenderedGridProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        legacyInsufficientRenderedGridProofCaseCount: 0,
        tenXMeanAndP95CaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        currentContractEvidenceComplete: true,
        googleSheetsTenXRequirementSatisfied: true,
        invalidReasons: [],
      },
    })
    expect(proof.cases[0]).toMatchObject({
      biligToGoogleSheetsMeanRatio: 0.05,
      biligToGoogleSheetsP95Ratio: 0.06,
      biligToMicrosoftExcelWebMeanRatio: 0.0625,
      biligToMicrosoftExcelWebP95Ratio: 0.06666666666666667,
      tenXMeanAndP95Metric: 'operationResponseMs',
      scenarioProof: {
        biligMeanMs: 5,
        biligP95Ms: 6,
        googleMeanMs: 100,
        googleP95Ms: 100,
        microsoftExcelWebMeanMs: 80,
        microsoftExcelWebP95Ms: 90,
        meanRatio: 0.05,
        p95Ratio: 0.06,
        microsoftExcelWebMeanRatio: 0.0625,
        microsoftExcelWebP95Ratio: 0.06666666666666667,
        screenshotProof: { captured: true, missingProducts: [] },
        pixelGridProof: { captured: true, missingProducts: [] },
      },
      postOperationFrameGuardrailPassed: true,
      passed: true,
    })
    expect(proof.cases.find((entry) => entry.workload === 'scroll-vertical')).toMatchObject({
      biligToGoogleSheetsScrollEventMeanRatio: 0.05,
      biligToGoogleSheetsScrollEventP95Ratio: 0.06,
      biligToMicrosoftExcelWebScrollEventMeanRatio: 0.0625,
      biligToMicrosoftExcelWebScrollEventP95Ratio: 0.06666666666666667,
      tenXMeanAndP95Metric: 'scrollEventResponseMs',
      scrollMovementGuardrailPassed: true,
      passed: true,
    })
  })

  it('rejects stale capture scenario proof before deriving same-corpus pass flags', () => {
    const capture = buildSameCorpusCapture()
    const weakCases = [...capture.cases]
    const firstCase = weakCases[0]
    weakCases[0] = Object.assign({}, firstCase, {
      scenarioProof: Object.assign({}, firstCase.scenarioProof, {
        pixelGridProof: Object.assign({}, firstCase.scenarioProof.pixelGridProof, {
          captured: true,
          products: firstCase.scenarioProof.pixelGridProof.products.map((productProof) =>
            productProof.product === 'bilig'
              ? Object.assign({}, productProof, {
                  evidence: ['mode=typegpu-v3', 'tilePaneCount=6', 'headerPaneCount=3'],
                })
              : productProof,
          ),
          missingProducts: [],
        }),
      }),
    })
    const weakCapture = withCaptureRunManifest({
      ...capture,
      cases: weakCases,
    })

    expect(() => buildSameCorpusProof(weakCapture)).toThrow('UI responsiveness same-corpus pixel grid proof is stale')
  })

  it('rejects stale raw same-corpus capture run manifests before deriving proof', () => {
    const capture = buildSameCorpusCapture()
    const forgedCapture: SameCorpusCapture = {
      ...capture,
      runManifest: {
        ...capture.runManifest,
        captureRunSignature: '0'.repeat(64),
        invalidReasons: ['hand-edited capture manifest'],
      },
    }

    expect(() => buildSameCorpusProof(forgedCapture)).toThrow('UI responsiveness same-corpus capture run manifest is stale')
  })

  it('keeps the same-corpus blocker for honestly reported weak Bilig pixel proof', () => {
    const capture = buildSameCorpusCapture()
    const weakCases = [...capture.cases]
    const firstCase = weakCases[0]
    const weakPixelGridProof = withProductPixelGridVerdicts({
      ...firstCase.scenarioProof.pixelGridProof,
      captured: false,
      products: firstCase.scenarioProof.pixelGridProof.products.map((productProof) =>
        productProof.product === 'bilig'
          ? Object.assign({}, productProof, {
              evidence: ['mode=typegpu-v3', 'tilePaneCount=6', 'headerPaneCount=3'],
            })
          : productProof,
      ),
      missingProducts: ['bilig'],
    })
    weakCases[0] = Object.assign({}, firstCase, {
      scenarioProof: Object.assign({}, firstCase.scenarioProof, {
        pixelGridProof: weakPixelGridProof,
      }),
    })
    const weakCapture = withCaptureRunManifest({
      ...capture,
      cases: weakCases,
    })

    const proof = buildSameCorpusProof(weakCapture)

    expect(proof.cases[0]).toMatchObject({
      passed: false,
      tenXMeanAndP95AgainstGoogleSheets: false,
      scenarioProof: {
        pixelGridProof: {
          captured: false,
          missingProducts: ['bilig'],
          productVerdicts: expect.arrayContaining([
            expect.objectContaining({
              product: 'bilig',
              evidenceStatus: 'legacy-insufficient',
              acceptedForCurrentScorecard: false,
            }),
          ]),
        },
      },
    })
    expect(proof.limitations).toContain(
      'Some same-corpus cases retain timing evidence but do not satisfy strict rendered-grid proof, so they cannot count toward Google Sheets 10x UI claims.',
    )
    expect(proof.runManifest).toMatchObject({
      strictRenderedGridProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length - 1,
      legacyInsufficientRenderedGridProofCaseCount: 1,
      currentContractEvidenceComplete: false,
      googleSheetsTenXRequirementSatisfied: false,
    })
    expect(proof.runManifest?.invalidReasons).toContain('strict rendered-grid proof covers 8/9 cases')
    expect(proof.runManifest?.invalidReasons).toContain('legacy-insufficient rendered-grid proof covers 1/9 cases')
  })

  it('keeps the same-corpus blocker when product source workbook fingerprints drift across workloads', () => {
    const capture = buildSameCorpusCapture()
    const driftedCases = [...capture.cases]
    const firstCase = driftedCases[0]
    driftedCases[0] = Object.assign({}, firstCase, {
      googleSheets: Object.assign({}, firstCase.googleSheets, {
        corpusVerification: Object.assign({}, firstCase.googleSheets.corpusVerification, {
          sourceWorkbookSha256: '3'.repeat(64),
        }),
      }),
    })
    const proof = buildSameCorpusProof(withCaptureRunManifest({ ...capture, cases: driftedCases }))

    expect(proof.tenXMeanAndP95CaseCount).toBe(requiredUiResponsivenessSameCorpusWorkloads.length)
    expect(proof.runManifest).toMatchObject({
      currentContractEvidenceComplete: false,
      googleSheetsTenXRequirementSatisfied: false,
    })
    expect(proof.runManifest?.invalidReasons).toContain('source workbook fingerprint must be stable for every required product')
  })

  it('rejects legacy operation-only same-corpus captures before generating proof', () => {
    expect(() =>
      buildSameCorpusProof(buildSameCorpusCapture({ includeScrollEventSamples: false, workloads: ['scroll-vertical'] })),
    ).toThrow('UI responsiveness same-corpus capture has too few scroll-event samples for same-corpus-wide-mixed-250k-scroll-vertical')
  })

  it('rejects captured same-corpus proof without scroll-event evidence', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const scrollCase = proof.cases.find((entry) => entry.workload === 'scroll-vertical')
    if (!scrollCase) {
      throw new Error('missing scroll-vertical fixture case')
    }
    const {
      scrollEventResponseMs: _biligScrollEventResponseMs,
      scrollMovementPx: _biligScrollMovementPx,
      ...biligWithoutScrollEvidence
    } = scrollCase.bilig
    const {
      scrollEventResponseMs: _googleSheetsScrollEventResponseMs,
      scrollMovementPx: _googleSheetsScrollMovementPx,
      ...googleSheetsWithoutScrollEvidence
    } = scrollCase.googleSheets
    const {
      scrollEventResponseMs: _microsoftExcelWebScrollEventResponseMs,
      scrollMovementPx: _microsoftExcelWebScrollMovementPx,
      ...microsoftExcelWebWithoutScrollEvidence
    } = scrollCase.microsoftExcelWeb
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry) =>
          entry.id === scrollCase.id
            ? Object.assign({}, scrollCase, {
                bilig: biligWithoutScrollEvidence,
                googleSheets: googleSheetsWithoutScrollEvidence,
                microsoftExcelWeb: microsoftExcelWebWithoutScrollEvidence,
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus proof is missing scroll-event evidence for same-corpus-wide-mixed-250k-scroll-vertical',
    )
  })

  it('allows same-corpus proof to clear the public-browser limitation blocker', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(
      hasUiResponsivenessSameCorpusTenXGap({
        ...scorecard,
        sameCorpusProof: buildSameCorpusProof(buildSameCorpusCapture()),
      }),
    ).toBe(false)
  })

  it('keeps the same-corpus blocker when required scroll evidence is missing', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(() => buildSameCorpusProof(buildSameCorpusCapture({ workloads: ['open-workbook'] }))).toThrow(
      'UI responsiveness same-corpus proof is missing required workload: select-cell',
    )
    expect(
      hasUiResponsivenessSameCorpusTenXGap({
        ...scorecard,
        sameCorpusProof: {
          ...scorecard.sameCorpusProof,
          captured: true,
          evidenceKind: 'same-corpus-browser-capture',
          requiredCaseCount: 1,
          tenXMeanAndP95CaseCount: 1,
          coveredCorpusCaseIds: ['wide-mixed-250k'],
          cases: [],
        },
      }),
    ).toBe(true)
  })

  it('rejects stale same-corpus pass flags and ratios', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) => (index === 0 ? Object.assign({}, entry, { biligToGoogleSheetsP95Ratio: 0.2 }) : entry)),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow('UI responsiveness same-corpus ratio is stale')
  })

  it('keeps the same-corpus blocker when scenario timing proof is hand-edited green', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const forgedScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) =>
          index === 0
            ? Object.assign({}, entry, {
                scenarioProof: Object.assign({}, entry.scenarioProof, {
                  meanRatio: 0.2,
                  p95Ratio: 0.2,
                }),
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(forgedScorecard)).toThrow(
      'UI responsiveness same-corpus scenario proof timing is stale',
    )
    expect(hasUiResponsivenessSameCorpusTenXGap(forgedScorecard)).toBe(true)
  })

  it('keeps the same-corpus blocker when TypeGPU pixel proof is hand-edited green', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const forgedScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) =>
          index === 0
            ? Object.assign({}, entry, {
                scenarioProof: Object.assign({}, entry.scenarioProof, {
                  pixelGridProof: Object.assign({}, entry.scenarioProof.pixelGridProof, {
                    captured: true,
                    missingProducts: [],
                    products: entry.scenarioProof.pixelGridProof.products.map((productProof) =>
                      productProof.product === 'bilig'
                        ? Object.assign({}, productProof, {
                            evidence: ['mode=typegpu-v3', 'pixelGridProofVersion=grid-pixels-v1'],
                          })
                        : productProof,
                    ),
                  }),
                }),
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(forgedScorecard)).toThrow(
      'UI responsiveness same-corpus pixel grid proof is stale',
    )
    expect(hasUiResponsivenessSameCorpusTenXGap(forgedScorecard)).toBe(true)
  })

  it('rejects stale same-corpus run manifests', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        runManifest: Object.assign({}, proof.runManifest, {
          strictRenderedGridProofCaseCount: 0,
          invalidReasons: ['hand-edited manifest'],
        }),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus run manifest is stale',
    )
  })

  it('rejects stale same-corpus visual proof required-product metadata', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) =>
          index === 0
            ? Object.assign({}, entry, {
                scenarioProof: Object.assign({}, entry.scenarioProof, {
                  screenshotProof: Object.assign({}, entry.scenarioProof.screenshotProof, {
                    requiredProducts: ['bilig', 'google-sheets', 'microsoft-excel-web'],
                  }),
                }),
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus screenshot proof is stale',
    )
  })
})

function buildSameCorpusCapture(
  args: {
    readonly includeScrollEventSamples?: boolean
    readonly workloads?: readonly UiResponsivenessSameCorpusWorkload[]
  } = {},
): SameCorpusCapture {
  const includeScrollEventSamples = args.includeScrollEventSamples ?? true
  const workloads = args.workloads ?? requiredUiResponsivenessSameCorpusWorkloads
  const cases = workloads.map((workload) => ({
    id: `same-corpus-wide-mixed-250k-${workload}`,
    corpusCaseId: 'wide-mixed-250k',
    materializedCells: 250000,
    workload,
    scenarioProof: sameCorpusScenarioProof(workload),
    bilig: {
      product: 'bilig' as const,
      source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
      operationResponseMsSamples: [4, 5, 6],
      postOperationFrameMsSamples: [8, 9, 10],
      ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
        ? { scrollEventResponseMsSamples: [4, 5, 6], scrollMovementPxSamples: [720, 720, 720] }
        : {}),
      corpusVerification: corpusVerification('bilig-benchmark-state', verifiedCells()),
      limitations: [],
    },
    googleSheets: {
      product: 'google-sheets' as const,
      source: 'https://docs.google.com/spreadsheets/d/example',
      operationResponseMsSamples: [100, 100, 100],
      postOperationFrameMsSamples: [14, 15, 16],
      ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
        ? { scrollEventResponseMsSamples: [100, 100, 100], scrollMovementPxSamples: [720, 720, 720] }
        : {}),
      corpusVerification: corpusVerification('google-sheets-xlsx-export', verifiedCells()),
      limitations: [],
    },
    microsoftExcelWeb: {
      product: 'microsoft-excel-web' as const,
      source: 'https://view.officeapps.live.com/op/view.aspx?src=example',
      operationResponseMsSamples: [75, 75, 90],
      postOperationFrameMsSamples: [14, 15, 16],
      ...(includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
        ? { scrollEventResponseMsSamples: [75, 75, 90], scrollMovementPxSamples: [720, 720, 720] }
        : {}),
      corpusVerification: corpusVerification('microsoft-excel-web-source-xlsx', verifiedCells()),
      limitations: [],
    },
  }))
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    runManifest: buildSameCorpusCaptureRunManifest(cases, 3),
    limitations: [],
    cases,
  }
}

function withCaptureRunManifest(capture: Omit<SameCorpusCapture, 'runManifest'>): SameCorpusCapture {
  return {
    ...capture,
    runManifest: buildSameCorpusCaptureRunManifest(capture.cases, capture.sampleCount),
  }
}

function sameCorpusScenarioProof(workload: UiResponsivenessSameCorpusWorkload) {
  const pixelGridProof = withProductPixelGridVerdicts({
    captured: true,
    requiredProducts: ['bilig', 'google-sheets'],
    products: [
      {
        product: 'bilig',
        captured: true,
        method: 'typegpu-visible-canvas',
        viewportPixelWidth: 1440,
        viewportPixelHeight: 900,
        evidence: [
          'gridCssWidth=720',
          'gridCssHeight=450',
          'devicePixelRatio=2',
          'expectedPixelWidth=1440',
          'expectedPixelHeight=900',
          'contractVersion=same-corpus-ui-v3',
          'gridAuthoritativeRevision=rev-3',
          'gridLocalRevision=rev-local-2',
          'gridProjectedRevision=rev-3',
          'fallbackMounted=false',
          'mode=typegpu-v3',
          'backendStatus=ready',
          'frameProofStatus=presented',
          'frameProofSignature=frame-current',
          'hasPresentedFrame=true',
          'hasPresentedVisibleFrame=true',
          'presentedFrameProofSignature=frame-current',
          'currentSceneOwnershipSignature=scene-current',
          'presentedSceneOwnershipSignature=scene-current',
          'currentContentSignature=content-current',
          'presentedContentSignature=content-current',
          'currentTextRunCount=12',
          'presentedTextRunCount=12',
          'currentTextSignature=text-current',
          'presentedTextSignature=text-current',
          'currentRectCount=88',
          'presentedRectCount=88',
          'currentRectSignature=rect-current',
          'presentedRectSignature=rect-current',
          'tilePaneCount=6',
          'headerPaneCount=3',
          'presentedTilePaneCount=6',
          'presentedHeaderPaneCount=3',
          'canvasPixelWidth=1440',
          'canvasPixelHeight=900',
          'canvasCoversViewport=true',
          'typeGpuAuthoritativeRevision=rev-3',
          'typeGpuLocalRevision=rev-local-2',
          'typeGpuProjectedRevision=rev-3',
          'visibleAuthoritativeRevision=rev-3',
          'visibleLocalRevision=rev-local-2',
          'visibleProjectedRevision=rev-3',
          'tileSceneRevision=scene-7',
          'visibleRenderRevision=scene-7',
          ...strictPixelGridEvidence(),
        ],
      },
      {
        product: 'google-sheets',
        captured: true,
        method: 'google-sheets-visible-grid',
        viewportPixelWidth: 1440,
        viewportPixelHeight: 900,
        evidence: strictPixelGridEvidence(),
      },
      {
        product: 'microsoft-excel-web',
        captured: true,
        method: 'excel-web-visible-grid',
        viewportPixelWidth: 1440,
        viewportPixelHeight: 900,
        evidence: strictPixelGridEvidence(),
      },
    ],
    missingProducts: [],
  })
  return {
    biligMeanMs: 5,
    biligP95Ms: 6,
    googleMeanMs: 100,
    googleP95Ms: 100,
    microsoftExcelWebMeanMs: 80,
    microsoftExcelWebP95Ms: 90,
    meanRatio: 0.05,
    p95Ratio: 0.06,
    microsoftExcelWebMeanRatio: 0.0625,
    microsoftExcelWebP95Ratio: 0.06666666666666667,
    screenshotProof: {
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      artifactPaths: [
        `tmp/same-corpus-wide-mixed-250k-${workload}/bilig-sample-1.png`,
        `tmp/same-corpus-wide-mixed-250k-${workload}/google-sheets-sample-1.png`,
        `tmp/same-corpus-wide-mixed-250k-${workload}/microsoft-excel-web-sample-1.png`,
      ],
      missingProducts: [],
    },
    pixelGridProof,
  }
}

function withProductPixelGridVerdicts(proof: Omit<SameCorpusPixelGridProof, 'productVerdicts'>): SameCorpusPixelGridProof {
  return {
    ...proof,
    productVerdicts: proof.products.map((entry: SameCorpusProductPixelGridProof) => validateSameCorpusProductPixelGridProof(entry)),
  }
}

function corpusVerification(
  method: 'bilig-benchmark-state' | 'google-sheets-xlsx-export' | 'microsoft-excel-web-source-xlsx',
  checkedCells: readonly { address: string; expected: string; actual: string }[],
) {
  const sourceWorkbookSha256 =
    method === 'bilig-benchmark-state'
      ? sameCorpusFixtureFingerprint.snapshotSha256
      : method === 'google-sheets-xlsx-export'
        ? googleSheetsSourceWorkbookSha256
        : microsoftExcelWebSourceWorkbookSha256
  return {
    verified: true,
    method,
    sheetName: 'WideGrid',
    materializedCells: 250000,
    corpusFingerprint: sameCorpusFixtureFingerprint,
    sourceWorkbookSha256,
    checkedCells,
  }
}

function verifiedCells() {
  return [
    { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
    { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
    { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
  ]
}

function strictPixelGridEvidence(): string[] {
  return [
    'pixelGridProofVersion=grid-pixels-v1',
    'pixelSampleSource=screenshot',
    'screenshotPixelWidth=1440',
    'screenshotPixelHeight=900',
    'nonBlankPixels=10000',
    'visibleGridLinePixels=4000',
    'verticalLineRuns=8',
    'horizontalLineRuns=16',
    'verticalLineCoverageBands=6',
    'horizontalLineCoverageBands=6',
    'largestVerticalLineGapPx=120',
    'largestHorizontalLineGapPx=24',
  ]
}
