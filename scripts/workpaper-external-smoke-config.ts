import { parseStrictBooleanEnvFlag } from './strict-env.js'

export function resolveKeepWorkpaperSmokeStage(env: { KEEP_WORKPAPER_SMOKE_STAGE?: string | undefined }): boolean {
  return parseStrictBooleanEnvFlag(env.KEEP_WORKPAPER_SMOKE_STAGE, 'KEEP_WORKPAPER_SMOKE_STAGE', false)
}
