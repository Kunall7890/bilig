import { describe, expect, it } from 'vitest'

import { isNpmDuplicateVersionPublishError } from '../runtime-package-set.ts'

describe('runtime package set helpers', () => {
  it('recognizes npm duplicate-version publish races as recoverable', () => {
    expect(
      isNpmDuplicateVersionPublishError(
        'npm error 403 Forbidden - PUT https://registry.npmjs.org/@bilig%2fprotocol - You cannot publish over the previously published versions: 0.51.2.',
      ),
    ).toBe(true)
    expect(isNpmDuplicateVersionPublishError('npm error 404 Not Found - PUT https://registry.npmjs.org/@bilig%2fworkbook')).toBe(false)
  })
})
