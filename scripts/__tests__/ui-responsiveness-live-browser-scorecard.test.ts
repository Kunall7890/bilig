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
  parseSameCorpusCapture,
  parseUiResponsivenessLiveBrowserCliArgs,
  parseUiResponsivenessLiveBrowserScorecard,
  sameCorpusScenarioCaseFields,
  validateSameCorpusCaptureArtifactMatchesScorecard,
  validateSameCorpusScreenshotArtifacts,
  validateUiResponsivenessLiveBrowserScorecard,
  type SameCorpusCapture,
  type SameCorpusCaptureMeasurement,
  type SameCorpusOperationResponseProof,
  type UiResponsivenessLiveBrowserScorecard,
  type UiResponsivenessSameCorpusProof,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'
import { buildSameCorpusFingerprint } from '../ui-responsiveness-same-corpus-fingerprint.ts'
import {
  buildCaptureScenarioProof,
  validateSameCorpusProductPixelGridProof,
  type SameCorpusPixelGridProof,
  type SameCorpusProductPixelGridProof,
  type SameCorpusProductVisualProof,
} from '../ui-responsiveness-same-corpus-proof.ts'
import {
  requiredUiResponsivenessSameCorpusWorkloads,
  uiSameCorpusWorkloadMutatesWorkbook,
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
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredProductCount: 2,
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      coveredCorpusCaseIds: ['wide-mixed-250k'],
    })
    expect(scorecard.sameCorpusProof.cases).toHaveLength(requiredUiResponsivenessSameCorpusWorkloads.length)
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBe(
      scorecard.sameCorpusProof.cases.filter((entry) => entry.tenXMeanAndP95AgainstGoogleSheets).length,
    )
    expect(scorecard.sameCorpusProof.tenXMeanAndP95CaseCount).toBe(0)
    expect(scorecard.sameCorpusProof.runManifest).toMatchObject({
      contractVersion: 'same-corpus-ui-v6',
      caseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      scenarioSummaryFieldCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      strictRenderedGridProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      visibleOperationResponseProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      biligAuthoritativeRenderProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      semanticUiProofCaseCount: 0,
      requiredMutationTargetProofCaseCount: 3,
      mutationTargetProofCaseCount: 0,
      legacyInsufficientRenderedGridProofCaseCount: 0,
      tenXMeanAndP95CaseCount: 0,
      currentContractEvidenceComplete: false,
      googleSheetsTenXRequirementSatisfied: false,
      captureRunSignature: expect.stringMatching(/^[a-f0-9]{64}$/u),
    })
    expect(scorecard.sameCorpusProof.runManifest?.capturedWorkloads).toEqual(requiredUiResponsivenessSameCorpusWorkloads)
    expect(scorecard.sameCorpusProof.runManifest?.invalidReasons).toContain('semantic UI proof covers 0/9 cases')
    expect(scorecard.sameCorpusProof.runManifest?.invalidReasons).toContain('mutation target proof covers 0/3 mutating cases')
    expect(scorecard.sameCorpusProof.runManifest?.invalidReasons).toContain('not every required workload is 10x against Google Sheets')
    expect(scorecard.sameCorpusProof.runManifest?.invalidReasons).not.toContain('Bilig authoritative render proof timing covers 0/9 cases')
    expect(scorecard.sameCorpusProof.limitations).toContain(
      'Caller must supply a Google Sheets URL for the same exported Bilig benchmark corpus.',
    )
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
    validateSameCorpusCaptureArtifactMatchesScorecard(scorecard)
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
        contractVersion: 'same-corpus-ui-v6',
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
        scenarioSummaryFieldCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        strictRenderedGridProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        visibleOperationResponseProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        biligAuthoritativeRenderProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        semanticUiProofCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        requiredMutationTargetProofCaseCount: 3,
        mutationTargetProofCaseCount: 3,
        legacyInsufficientRenderedGridProofCaseCount: 0,
        tenXMeanAndP95CaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
        currentContractEvidenceComplete: true,
        googleSheetsTenXRequirementSatisfied: true,
        invalidReasons: [],
      },
    })
    expect(proof.cases[0]).toMatchObject({
      biligMeanMs: 5,
      biligP95Ms: 6,
      googleMeanMs: 100,
      googleP95Ms: 100,
      microsoftExcelWebMeanMs: 80,
      microsoftExcelWebP95Ms: 90,
      meanRatio: 20,
      p95Ratio: 16.666666666666668,
      microsoftExcelWebMeanRatio: 16,
      microsoftExcelWebP95Ratio: 15,
      screenshotProof: { captured: true, missingProducts: [] },
      pixelGridProof: { captured: true, missingProducts: [] },
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
        meanRatio: 20,
        p95Ratio: 16.666666666666668,
        microsoftExcelWebMeanRatio: 16,
        microsoftExcelWebP95Ratio: 15,
        screenshotProof: { captured: true, missingProducts: [] },
        pixelGridProof: { captured: true, missingProducts: [] },
      },
      postOperationFrameGuardrailPassed: true,
      authoritativeRenderProofGuardrailPassed: true,
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

  it('rejects stale first-class same-corpus capture scenario fields before deriving proof', () => {
    const capture = buildSameCorpusCapture()
    const staleCases = [...capture.cases]
    staleCases[0] = Object.assign({}, staleCases[0], {
      biligMeanMs: 999,
    })
    const staleCapture = withCaptureRunManifest({
      ...capture,
      cases: staleCases,
    })

    expect(() => buildSameCorpusProof(staleCapture)).toThrow('UI responsiveness same-corpus capture scenario summary fields are stale')
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
    const scenarioProof = Object.assign({}, firstCase.scenarioProof, {
      pixelGridProof: weakPixelGridProof,
    })
    weakCases[0] = Object.assign({}, firstCase, {
      ...sameCorpusScenarioCaseFields(scenarioProof),
      scenarioProof,
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

  it('rejects stale first-class same-corpus scorecard scenario fields', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        cases: proof.cases.map((entry, index) => (index === 0 ? Object.assign({}, entry, { googleMeanMs: 999 }) : entry)),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus scorecard scenario summary fields are stale',
    )
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
                  meanRatio: 10,
                  p95Ratio: 10,
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

  it('rejects captured same-corpus proof without a capture run signature', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: {
        ...proof,
        runManifest: Object.assign({}, proof.runManifest, {
          captureRunSignature: null,
        }),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus proof run manifest must bind to a captureRunSignature',
    )
  })

  it('rejects captured same-corpus proof when screenshot artifacts are missing', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-same-corpus-missing-artifacts-`)

    expect(() => validateSameCorpusScreenshotArtifacts(proof, { rootDir })).toThrow(
      'UI responsiveness same-corpus screenshot artifact is missing',
    )
  })

  it('requires captured same-corpus screenshot artifacts to be tracked for checked-in proof', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-same-corpus-tracked-artifacts-`)
    writeSameCorpusScreenshotArtifacts(rootDir, proof)
    const artifactPaths = sameCorpusScreenshotArtifactPaths(proof)

    expect(() =>
      validateSameCorpusScreenshotArtifacts(proof, {
        requireGitTracked: true,
        rootDir,
        trackedArtifactPaths: artifactPaths.slice(1),
      }),
    ).toThrow('UI responsiveness same-corpus screenshot artifact is not tracked by git')
    expect(() =>
      validateSameCorpusScreenshotArtifacts(proof, {
        requireGitTracked: true,
        rootDir,
        trackedArtifactPaths: artifactPaths,
      }),
    ).not.toThrow()
  })

  it('requires semantic mutation target screenshot artifacts to exist', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-same-corpus-mutation-artifacts-`)
    for (const artifactPath of sameCorpusScenarioScreenshotArtifactPaths(proof)) {
      writeSameCorpusScreenshotArtifact(rootDir, artifactPath)
    }

    expect(() => validateSameCorpusScreenshotArtifacts(proof, { rootDir })).toThrow(
      'UI responsiveness same-corpus screenshot artifact is missing',
    )
  })

  it('requires semantic mutation target screenshot artifacts to be tracked for checked-in proof', () => {
    const proof = buildSameCorpusProof(buildSameCorpusCapture())
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-same-corpus-tracked-mutation-artifacts-`)
    writeSameCorpusScreenshotArtifacts(rootDir, proof)
    const artifactPaths = sameCorpusScreenshotArtifactPaths(proof)
    const mutationArtifactPath = sameCorpusMutationTargetScreenshotArtifactPaths(proof)[0]

    expect(mutationArtifactPath).toContain('/mutation-target/')
    expect(() =>
      validateSameCorpusScreenshotArtifacts(proof, {
        requireGitTracked: true,
        rootDir,
        trackedArtifactPaths: artifactPaths.filter((artifactPath) => artifactPath !== mutationArtifactPath),
      }),
    ).toThrow('UI responsiveness same-corpus screenshot artifact is not tracked by git')
  })

  it('requires checked-in same-corpus scorecard proof to match the capture artifact', () => {
    const capture = buildSameCorpusCapture()
    const rootDir = mkdtempSync(`${tmpdir()}/bilig-same-corpus-capture-artifact-`)
    const capturePath = '.cache/ui-responsiveness/same-corpus-capture.json'
    mkdirSync(dirname(resolve(rootDir, capturePath)), { recursive: true })
    writeFileSync(resolve(rootDir, capturePath), `${JSON.stringify(capture, null, 2)}\n`)
    const proof = buildSameCorpusProof(parseSameCorpusCapture(readJsonObject(resolve(rootDir, capturePath))))
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const matchingScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      sameCorpusProof: proof,
    }
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...matchingScorecard,
      sameCorpusProof: {
        ...proof,
        tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount + 1,
      },
    }
    const staleSignatureScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...matchingScorecard,
      sameCorpusProof: {
        ...proof,
        runManifest: {
          ...proof.runManifest,
          captureRunSignature: 'f'.repeat(64),
        },
      },
    }

    expect(() => validateSameCorpusCaptureArtifactMatchesScorecard(matchingScorecard, { rootDir, capturePath })).not.toThrow()
    expect(() => validateSameCorpusCaptureArtifactMatchesScorecard(staleScorecard, { rootDir, capturePath })).toThrow(
      'UI responsiveness same-corpus scorecard proof does not match capture artifact',
    )
    expect(() => validateSameCorpusCaptureArtifactMatchesScorecard(staleSignatureScorecard, { rootDir, capturePath })).toThrow(
      'UI responsiveness same-corpus scorecard proof does not match capture artifact',
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

  it('rejects same-corpus screenshot artifacts that are not tied to the exact scenario', () => {
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
                    artifactPaths: entry.scenarioProof.screenshotProof.artifactPaths.map((artifactPath) =>
                      artifactPath.replace(entry.id, 'same-corpus-wide-mixed-250k-other-case'),
                    ),
                  }),
                }),
              })
            : entry,
        ),
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness same-corpus screenshot artifact path is not tied to scenario',
    )
  })
})

function writeSameCorpusScreenshotArtifacts(rootDir: string, proof: UiResponsivenessSameCorpusProof): void {
  for (const artifactPath of sameCorpusScreenshotArtifactPaths(proof)) {
    writeSameCorpusScreenshotArtifact(rootDir, artifactPath)
  }
}

function writeSameCorpusScreenshotArtifact(rootDir: string, artifactPath: string): void {
  const absolutePath = resolve(rootDir, artifactPath)
  mkdirSync(dirname(absolutePath), { recursive: true })
  writeFileSync(absolutePath, 'png')
}

function sameCorpusScreenshotArtifactPaths(proof: UiResponsivenessSameCorpusProof): string[] {
  return [
    ...new Set([...sameCorpusScenarioScreenshotArtifactPaths(proof), ...sameCorpusMutationTargetScreenshotArtifactPaths(proof)]),
  ].toSorted()
}

function sameCorpusScenarioScreenshotArtifactPaths(proof: UiResponsivenessSameCorpusProof): string[] {
  return [...new Set(proof.cases.flatMap((entry) => [...entry.scenarioProof.screenshotProof.artifactPaths]))].toSorted()
}

function sameCorpusMutationTargetScreenshotArtifactPaths(proof: UiResponsivenessSameCorpusProof): string[] {
  return [
    ...new Set(
      proof.cases.flatMap((entry) =>
        entry.scenarioProof.semanticUiProof.products.flatMap((productProof) =>
          productProof.mutationTargetProofs.flatMap((mutationProof) => mutationProof.screenshotPath ?? []),
        ),
      ),
    ),
  ].toSorted()
}

function buildSameCorpusCapture(
  args: {
    readonly includeScrollEventSamples?: boolean
    readonly workloads?: readonly UiResponsivenessSameCorpusWorkload[]
  } = {},
): SameCorpusCapture {
  const includeScrollEventSamples = args.includeScrollEventSamples ?? true
  const workloads = args.workloads ?? requiredUiResponsivenessSameCorpusWorkloads
  const cases = workloads.map((workload) => {
    const bilig = sameCorpusCaptureMeasurementFixture('bilig', workload, includeScrollEventSamples)
    const googleSheets = sameCorpusCaptureMeasurementFixture('google-sheets', workload, includeScrollEventSamples)
    const microsoftExcelWeb = sameCorpusCaptureMeasurementFixture('microsoft-excel-web', workload, includeScrollEventSamples)
    const scenarioProof = sameCorpusScenarioProof(workload, bilig, googleSheets, microsoftExcelWeb)
    return Object.assign(
      {
        id: `same-corpus-wide-mixed-250k-${workload}`,
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250000,
        workload,
      },
      sameCorpusScenarioCaseFields(scenarioProof),
      {
        scenarioProof,
        bilig,
        googleSheets,
        microsoftExcelWeb,
      },
    )
  })
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

function sameCorpusCaptureMeasurementFixture(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  workload: UiResponsivenessSameCorpusWorkload,
  includeScrollEventSamples: boolean,
): SameCorpusCaptureMeasurement {
  const requiresScrollEventSamples = includeScrollEventSamples && uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)
  const operationResponseMsSamples = product === 'bilig' ? [4, 5, 6] : product === 'google-sheets' ? [100, 100, 100] : [75, 75, 90]
  const source =
    product === 'bilig'
      ? 'e2e/tests/web-shell-scroll-performance.pw.ts'
      : product === 'google-sheets'
        ? 'https://docs.google.com/spreadsheets/d/example'
        : 'https://view.officeapps.live.com/op/view.aspx?src=example'
  const method =
    product === 'bilig'
      ? 'bilig-benchmark-state'
      : product === 'google-sheets'
        ? 'google-sheets-xlsx-export'
        : 'microsoft-excel-web-source-xlsx'
  return {
    product,
    source,
    operationResponseMsSamples,
    operationResponseProofs: sameCorpusOperationResponseProofSamples(workload),
    ...(product === 'bilig' ? { authoritativeRenderProofMsSamples: [9, 10, 11] } : {}),
    postOperationFrameMsSamples: product === 'bilig' ? [8, 9, 10] : [14, 15, 16],
    ...(requiresScrollEventSamples
      ? { scrollEventResponseMsSamples: operationResponseMsSamples, scrollMovementPxSamples: [720, 720, 720] }
      : {}),
    ...(product === 'bilig' ? { biligRuntimeProof: biligRuntimeProofFixture(source) } : {}),
    corpusVerification: corpusVerification(method, verifiedCells()),
    limitations: [],
  }
}

function sameCorpusOperationResponseProofSamples(workload: UiResponsivenessSameCorpusWorkload): SameCorpusOperationResponseProof[] {
  const proof = expectedSameCorpusOperationResponseProof(workload)
  return [proof, proof, proof]
}

function expectedSameCorpusOperationResponseProof(workload: UiResponsivenessSameCorpusWorkload): SameCorpusOperationResponseProof {
  if (workload === 'open-workbook') {
    return 'load-to-ready'
  }
  if (uiSameCorpusWorkloadRequiresScrollEventEvidence(workload)) {
    return 'visible-scroll-movement'
  }
  return 'visible-non-scroll-response'
}

function biligRuntimeProofFixture(source: string) {
  return {
    product: 'bilig' as const,
    source,
    verificationMethod: 'window.__biligRuntimeBuild' as const,
    requiredBuildKind: 'production' as const,
    actualBuildKind: 'production' as const,
    mode: 'production',
    dev: false,
    prod: true,
    remoteSyncEnabled: false,
    entryRoute: '/?benchmarkCorpus=wide-mixed-250k&persist=0',
    sampleCount: 3,
    verified: true,
    samples: [0, 1, 2].map((sampleIndex) => ({
      sampleIndex,
      present: true,
      app: '@bilig/web',
      buildKind: 'production' as const,
      mode: 'production',
      dev: false,
      prod: true,
      remoteSyncEnabled: false,
      entryRoute: '/?benchmarkCorpus=wide-mixed-250k&persist=0',
    })),
  }
}

function sameCorpusScenarioProof(
  workload: UiResponsivenessSameCorpusWorkload,
  bilig: SameCorpusCaptureMeasurement,
  googleSheets: SameCorpusCaptureMeasurement,
  microsoftExcelWeb: SameCorpusCaptureMeasurement,
) {
  const visualProofs: SameCorpusProductVisualProof[] = [
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
        'contractVersion=same-corpus-ui-v6',
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
        'currentSceneEpochSignature=epoch-current',
        'currentSceneOwnershipSignature=scene-current',
        'presentedSceneEpochSignature=epoch-current',
        'presentedSceneOwnershipSignature=scene-current',
        'currentWorkbookRevision=workbook-current',
        'presentedWorkbookRevision=workbook-current',
        'currentSemanticMutationRevision=semantic-current',
        'presentedSemanticMutationRevision=semantic-current',
        'currentViewportRevision=viewport-current',
        'presentedViewportRevision=viewport-current',
        'currentSelectionRevision=selection-current',
        'presentedSelectionRevision=selection-current',
        'currentFillHandleRevision=fill-current',
        'presentedFillHandleRevision=fill-current',
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
  ].map((pixelGridProof) => ({
    product: pixelGridProof.product,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/${pixelGridProof.product}-sample-1.png`,
    screenshotCaptured: true,
    pixelGridProof,
    semanticUiProof: semanticUiProofFixture(
      pixelGridProof.product,
      pixelGridProof.product === 'bilig'
        ? bilig.corpusVerification
        : pixelGridProof.product === 'google-sheets'
          ? googleSheets.corpusVerification
          : microsoftExcelWeb.corpusVerification,
      workload,
    ),
  }))
  return buildCaptureScenarioProof({ bilig, googleSheets, microsoftExcelWeb, visualProofs, workload })
}

function semanticUiProofFixture(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  verification: SameCorpusCaptureMeasurement['corpusVerification'],
  workload: UiResponsivenessSameCorpusWorkload,
) {
  return {
    product,
    captured: true,
    method:
      product === 'bilig'
        ? ('bilig-visible-semantic-readback' as const)
        : product === 'google-sheets'
          ? ('google-sheets-visible-semantic-readback' as const)
          : ('excel-web-visible-semantic-readback' as const),
    sheetName: verification.sheetName,
    sheetId: product === 'bilig' ? 'sheet-wide-grid' : null,
    selectedRange: 'A1',
    checkedCells: verification.checkedCells,
    authoritativeRenderRevision: product === 'bilig' ? 'rev-3' : null,
    visibleRenderRevision: product === 'bilig' ? 'scene-7' : null,
    screenshotSha256: 'a'.repeat(64),
    mutationTargetProofs: sameCorpusMutationTargetProofs(product, workload),
    evidence: [
      `sheetName=${verification.sheetName}`,
      'selectedRange=A1',
      `checkedCellCount=${String(verification.checkedCells.length)}`,
      'screenshotSha256=' + 'a'.repeat(64),
      ...(product === 'bilig' ? ['authoritativeRenderRevision=rev-3', 'visibleRenderRevision=scene-7'] : []),
    ],
  }
}

function sameCorpusMutationTargetProofs(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  workload: UiResponsivenessSameCorpusWorkload,
) {
  if (!uiSameCorpusWorkloadMutatesWorkbook(workload)) {
    return []
  }
  return [0, 1, 2].map((sampleIndex) => ({
    sampleIndex,
    workload,
    intendedOperation: workload,
    intendedPayload: sameCorpusMutationTargetIntendedPayload(product, workload, sampleIndex),
    sheetName: 'WideGrid',
    targetRange: 'A1',
    before: sameCorpusMutationReadback(product, workload, 'before', sampleIndex),
    after: sameCorpusMutationReadback(product, workload, 'after', sampleIndex),
    restored: sameCorpusMutationReadback(product, workload, 'before', sampleIndex),
    visibleAfter: sameCorpusVisibleMutationReadback(product, workload, 'after', sampleIndex),
    visibleRestored: sameCorpusVisibleMutationReadback(product, workload, 'before', sampleIndex),
    authoritativeReadbackRevision: `authoritative-readback-${sampleIndex + 1}`,
    visibleRenderRevision: `visible-render-${sampleIndex + 1}`,
    screenshotPath: `tmp/same-corpus-wide-mixed-250k-${workload}/mutation-target/${product}-sample-${sampleIndex + 1}-after.png`,
    screenshotSha256: 'a'.repeat(64),
    undoRestoreStatus: 'verified' as const,
  }))
}

function sameCorpusMutationTargetIntendedPayload(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  workload: UiResponsivenessSameCorpusWorkload,
  sampleIndex: number,
) {
  if (workload === 'formula-edit') {
    return { kind: 'formula' as const, formula: `=${sampleIndex + 1}+1` }
  }
  if (workload === 'fill-format-change') {
    const labels = ['light cornflower blue 3', 'theme green', 'light cornflower blue 2'] as const
    return { kind: 'fill-color' as const, swatchLabel: labels[sampleIndex % labels.length] }
  }
  return { kind: 'cell-value' as const, value: `${product}-same-corpus-${sampleIndex + 1}` }
}

function sameCorpusVisibleMutationReadback(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  workload: UiResponsivenessSameCorpusWorkload,
  phase: 'before' | 'after',
  sampleIndex: number,
) {
  return {
    ...sameCorpusMutationReadback(product, workload, phase, sampleIndex),
    source: 'visible-formula-bar' as const,
  }
}

function sameCorpusMutationReadback(
  product: 'bilig' | 'google-sheets' | 'microsoft-excel-web',
  workload: UiResponsivenessSameCorpusWorkload,
  phase: 'before' | 'after',
  sampleIndex: number,
) {
  const after = phase === 'after'
  const source = product === 'bilig' ? ('bilig-authoritative-range' as const) : ('visible-formula-bar' as const)
  if (workload === 'formula-edit') {
    return {
      value: after ? String(sampleIndex + 2) : 'metric-1',
      formula: after ? `=${String(sampleIndex + 1)}+1` : null,
      fillColor: null,
      visibleText: after ? String(sampleIndex + 2) : 'metric-1',
      source,
    }
  }
  if (workload === 'fill-format-change') {
    return {
      value: 'metric-1',
      formula: null,
      fillColor: after ? '#c9daf8' : null,
      visibleText: 'metric-1',
      source,
    }
  }
  return {
    value: after ? `${product}-same-corpus-${String(sampleIndex + 1)}` : 'metric-1',
    formula: null,
    fillColor: null,
    visibleText: after ? `${product}-same-corpus-${String(sampleIndex + 1)}` : 'metric-1',
    source,
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
