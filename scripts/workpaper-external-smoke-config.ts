export function resolveKeepWorkpaperSmokeStage(env: { KEEP_WORKPAPER_SMOKE_STAGE?: string | undefined }): boolean {
  return parseStrictBooleanFlag(env.KEEP_WORKPAPER_SMOKE_STAGE, 'KEEP_WORKPAPER_SMOKE_STAGE', false)
}

function parseStrictBooleanFlag(value: string | undefined, name: string, fallback: boolean): boolean {
  if (value === undefined || value.length === 0) {
    return fallback
  }
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  throw new Error(`${name} must be "1", "true", "0", or "false" when set, got ${value}`)
}
