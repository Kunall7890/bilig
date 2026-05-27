import { readFileSync } from 'node:fs'
import { createServer, type Server } from 'node:net'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

import {
  resolveServedBiligProductionPort,
  sameCorpusProductionBuildEnv,
  sameCorpusProductionProofApiEnvFlag,
} from '../capture-ui-responsiveness-same-corpus.ts'

describe('same-corpus production Bilig runtime', () => {
  it('builds served production captures with the benchmark proof API explicitly enabled', () => {
    expect(sameCorpusProductionBuildEnv({ EXISTING_FLAG: 'kept' })).toEqual({
      EXISTING_FLAG: 'kept',
      [sameCorpusProductionProofApiEnvFlag]: '1',
    })
  })

  it('does not shadow the global process object while starting the preview server', () => {
    const source = readFileSync(resolve(import.meta.dirname, '../capture-ui-responsiveness-same-corpus.ts'), 'utf8')

    expect(source).not.toContain('const process = Bun.spawn')
    expect(source).toContain('sameCorpusProductionBuildEnv(process.env)')
  })

  it('moves default served captures off an occupied preview port', async () => {
    const occupied = await listenOnEphemeralLoopbackPort()
    try {
      const address = occupied.address()
      const port = typeof address === 'object' && address ? address.port : 0

      await expect(
        resolveServedBiligProductionPort({
          biligProductionHost: '127.0.0.1',
          biligProductionPort: port,
          biligProductionPortSource: 'default',
        }),
      ).resolves.not.toBe(port)
    } finally {
      await closeServer(occupied)
    }
  })

  it('keeps explicit served capture preview ports strict', async () => {
    const occupied = await listenOnEphemeralLoopbackPort()
    try {
      const address = occupied.address()
      const port = typeof address === 'object' && address ? address.port : 0

      await expect(
        resolveServedBiligProductionPort({
          biligProductionHost: '127.0.0.1',
          biligProductionPort: port,
          biligProductionPortSource: 'explicit',
        }),
      ).resolves.toBe(port)
    } finally {
      await closeServer(occupied)
    }
  })
})

function listenOnEphemeralLoopbackPort(): Promise<Server> {
  return new Promise((resolveServer, reject) => {
    const server = createServer()
    server.once('error', reject)
    server.listen({ host: '127.0.0.1', port: 0 }, () => {
      server.off('error', reject)
      resolveServer(server)
    })
  })
}

function closeServer(server: Server): Promise<void> {
  return new Promise((resolveClose, reject) => {
    server.close((error) => {
      if (error) {
        reject(error)
        return
      }
      resolveClose()
    })
  })
}
