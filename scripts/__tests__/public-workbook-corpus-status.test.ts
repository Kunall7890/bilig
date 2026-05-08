import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import { writePublicWorkbookCorpusCheck } from '../public-workbook-corpus-status.ts'

describe('public workbook corpus status checks', () => {
  it('points target-incomplete checks at the safe resume-plan command', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-status-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const scorecardPath = join(dir, 'scorecard.json')
    const verifyCheckpointPath = join(dir, 'verification-checkpoint.json')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-08T11:45:00.000Z', 2), null, 2)}\n`)

    expect(() =>
      writePublicWorkbookCorpusCheck({
        cacheDir,
        manifestPath,
        requireTarget: true,
        scorecardPath,
        skipManifestCheck: false,
        verifyCheckpointPath,
      }),
    ).toThrowError(
      new RegExp(
        [
          'Public workbook corpus target incomplete:',
          'cached artifacts below target: 0/2',
          'next command: pnpm public-workbook-corpus:resume-plan:check -- --manifest',
          'manifest\\.json',
          '--cache-dir',
          'cache',
          '--scorecard',
          'scorecard\\.json',
          '--verify-checkpoint',
          'verification-checkpoint\\.json',
        ].join('.*'),
        'su',
      ),
    )
  })
})
