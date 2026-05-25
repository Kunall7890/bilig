import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { buildPublicClaimCheckReport, collectPublicClaimFiles, findBroadGoogleSheetsTenXClaims } from '../check-public-claims.ts'
import { rootDir } from '../bilig-dominance-scorecard-input.ts'

const tempRoots: string[] = []

afterEach(() => {
  for (const root of tempRoots.splice(0)) {
    rmSync(root, { recursive: true, force: true })
  }
})

describe('public claim check', () => {
  it('finds broad Google Sheets 10x claims', () => {
    expect(findBroadGoogleSheetsTenXClaims('Bilig is 10x faster than Google Sheets for every workbook.', 'README.md')).toEqual([
      {
        path: 'README.md',
        line: 1,
        column: 10,
        match: '10x faster than Google Sheets',
        text: 'Bilig is 10x faster than Google Sheets for every workbook.',
      },
    ])

    expect(findBroadGoogleSheetsTenXClaims('Bilig beats Google Sheets by 10x on UI.', 'docs/index.html')).toHaveLength(1)
    expect(findBroadGoogleSheetsTenXClaims('No blanket 10x claim is allowed yet.', 'README.md')).toEqual([])
  })

  it('blocks public broad claims while blanket 10x claims are disallowed', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Bilig is 10x better than Google Sheets.',
      'docs/index.html': '<h1>Scoped benchmark evidence</h1>',
    })

    const report = buildPublicClaimCheckReport({
      repoRoot,
      scorecard: { claimPolicy: { blanketTenXClaimAllowed: false } },
    })

    expect(report.violations).toHaveLength(1)
    expect(report.violations[0]).toMatchObject({
      path: 'README.md',
      match: '10x better than Google Sheets',
    })
  })

  it('allows broad claims once the full dominance scorecard gate allows them', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Bilig is 10x faster than Google Sheets.',
    })

    const report = buildPublicClaimCheckReport({
      repoRoot,
      scorecard: passingClaimScorecard(),
    })

    expect(report.violations).toEqual([])
  })

  it('blocks broad claims when only the blanket claim boolean is forged', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Bilig is 10x faster than Google Sheets.',
    })

    const report = buildPublicClaimCheckReport({
      repoRoot,
      scorecard: { claimPolicy: { blanketTenXClaimAllowed: true } },
    })

    expect(report.blanketTenXClaimAllowed).toBe(false)
    expect(report.violations).toHaveLength(1)
  })

  it('scans public docs while excluding internal planning docs', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Root readme',
      'docs/index.html': '<main>Public site</main>',
      'docs/public-api.md': '# Public API',
      'docs/workbook-view-platform-10x-production-plan-2026-04-29.md': '# Internal plan',
      'packages/headless/README.md': '# Headless',
    })

    expect(collectPublicClaimFiles(repoRoot)).toEqual(['README.md', 'docs/index.html', 'docs/public-api.md', 'packages/headless/README.md'])
  })

  it('passes for the current checked-in public surfaces', () => {
    expect(buildPublicClaimCheckReport({ repoRoot: rootDir }).violations).toEqual([])
  })
})

function makeTempRepo(files: Record<string, string>): string {
  const repoRoot = mkdtempSync(join(tmpdir(), 'bilig-public-claims-'))
  tempRoots.push(repoRoot)
  for (const [repoPath, source] of Object.entries(files)) {
    const path = join(repoRoot, repoPath)
    mkdirSync(dirname(path), { recursive: true })
    writeFileSync(path, source)
  }
  return repoRoot
}

function passingClaimScorecard(): unknown {
  return {
    goalStatus: 'achieved',
    claimPolicy: {
      blanketTenXClaimAllowed: true,
      unmetRequirements: [],
    },
    completionAudit: {
      allCriteriaPassed: true,
      unmetRequirements: [],
      criteria: [],
    },
    overallGoogleSheets10xStatus: {
      passed: true,
      status: 'passed',
      unmetRequirements: [],
      categories: [
        { id: 'recalculation-speed', passed: true, gaps: [] },
        { id: 'structural-edit-performance', passed: true, gaps: [] },
        { id: 'large-workbook-scale', passed: true, gaps: [] },
        { id: 'ui-responsiveness', passed: true, gaps: [] },
      ],
    },
  }
}
