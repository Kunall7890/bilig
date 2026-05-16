import { parseStrictBooleanEnvFlag } from './strict-env.js'

export type CiProfile = 'fast' | 'full'

export function resolveCiProfile(env: Readonly<Record<string, string | undefined>>): CiProfile {
  const value = env['BILIG_CI_PROFILE']
  if (value === undefined || value === 'fast') {
    return 'fast'
  }
  if (value === 'full') {
    return 'full'
  }

  throw new Error(`BILIG_CI_PROFILE must be "fast" or "full", got ${value}`)
}

export function resolveCiSkipBrowserGates(env: Readonly<Record<string, string | undefined>>): boolean {
  return parseStrictBooleanEnvFlag(env['BILIG_CI_SKIP_BROWSER'], 'BILIG_CI_SKIP_BROWSER', false)
}
