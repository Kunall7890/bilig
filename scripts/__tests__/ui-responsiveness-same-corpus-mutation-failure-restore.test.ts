import { describe, expect, it } from 'vitest'

import { withSameCorpusMutationFailureRestore } from '../ui-responsiveness-same-corpus-mutation-failure-restore.ts'

describe('same-corpus mutating sample failure restore', () => {
  it('best-effort restores and reselects a mutating target before rethrowing the original failure', async () => {
    const calls: string[] = []
    const failure = new Error('mutation operation failed after write')

    await expect(
      withSameCorpusMutationFailureRestore({
        workload: 'edit-visible-cell',
        run: async () => {
          calls.push('run')
          throw failure
        },
        restore: async () => {
          calls.push('restore')
        },
        reselectTarget: async () => {
          calls.push('reselect')
        },
      }),
    ).rejects.toBe(failure)

    expect(calls).toEqual(['run', 'restore', 'reselect'])
  })

  it('preserves the original mutating failure when best-effort cleanup also fails', async () => {
    const calls: string[] = []
    const failure = new Error('target proof setup failed')

    await expect(
      withSameCorpusMutationFailureRestore({
        workload: 'fill-format-change',
        run: async () => {
          calls.push('run')
          throw failure
        },
        restore: async () => {
          calls.push('restore')
          throw new Error('undo failed')
        },
        reselectTarget: async () => {
          calls.push('reselect')
          throw new Error('reselect failed')
        },
      }),
    ).rejects.toBe(failure)

    expect(calls).toEqual(['run', 'restore', 'reselect'])
  })

  it('does not send an undo for non-mutating workloads', async () => {
    const calls: string[] = []
    const failure = new Error('selection failed')

    await expect(
      withSameCorpusMutationFailureRestore({
        workload: 'select-cell',
        run: async () => {
          calls.push('run')
          throw failure
        },
        restore: async () => {
          calls.push('restore')
        },
        reselectTarget: async () => {
          calls.push('reselect')
        },
      }),
    ).rejects.toBe(failure)

    expect(calls).toEqual(['run'])
  })

  it('does not run cleanup after a successful mutating sample', async () => {
    const calls: string[] = []

    await expect(
      withSameCorpusMutationFailureRestore({
        workload: 'formula-edit',
        run: async () => {
          calls.push('run')
          return 'ok'
        },
        restore: async () => {
          calls.push('restore')
        },
        reselectTarget: async () => {
          calls.push('reselect')
        },
      }),
    ).resolves.toBe('ok')

    expect(calls).toEqual(['run'])
  })
})
