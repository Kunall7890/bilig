import type { LargeSimpleXlsxImportPhaseRecorder } from './xlsx-large-simple-import-telemetry.js'
import type { LargeSimpleXlsxImportOptions } from './xlsx-large-simple-import-types.js'
import { readLazyXlsxZipSourceByteLength, replaceLazyXlsxZipSource, type XlsxZipEntries } from './xlsx-zip.js'

export function replaceLargeSimpleZipSourceForImport(
  zip: XlsxZipEntries,
  options: LargeSimpleXlsxImportOptions,
  phaseRecorder: LargeSimpleXlsxImportPhaseRecorder,
): boolean {
  if (options.releaseZipSource !== true || !options.replacementZipSource) {
    return false
  }
  const zipSourceReleaseStart = phaseRecorder.start()
  const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
  const zipSourceReplaced = replaceLazyXlsxZipSource(zip, options.replacementZipSource)
  const ownedSourceReleaseEvidence = zipSourceReplaced ? options.releaseOwnedSourceBytes?.() : undefined
  phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
    ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
    ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
    ...ownedSourceReleaseEvidence,
  })
  return Boolean(ownedSourceReleaseEvidence)
}
