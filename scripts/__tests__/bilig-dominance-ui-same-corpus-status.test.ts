import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { readUiSameCorpusCaptureArtifactStatus } from '../bilig-dominance-ui-same-corpus-status.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'

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
        readinessErrors: ['Expected runManifest to be an object'],
        runManifestInvalidReasons: [],
      })
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
