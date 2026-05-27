import { describe, expect, it } from 'vitest'

import { sameCorpusProductionBuildEnv, sameCorpusProductionProofApiEnvFlag } from '../capture-ui-responsiveness-same-corpus.ts'

describe('same-corpus production Bilig runtime', () => {
  it('builds served production captures with the benchmark proof API explicitly enabled', () => {
    expect(sameCorpusProductionBuildEnv({ EXISTING_FLAG: 'kept' })).toEqual({
      EXISTING_FLAG: 'kept',
      [sameCorpusProductionProofApiEnvFlag]: '1',
    })
  })
})
