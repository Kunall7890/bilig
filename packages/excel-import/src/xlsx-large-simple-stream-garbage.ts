import { collectLargeSimpleImportGarbage } from './xlsx-large-simple-garbage.js'
import { forEachInflatedXlsxZipEntryChunk, type XlsxZipEntries } from './xlsx-zip.js'

const defaultGarbageCollectionIntervalBytes = 128 * 1024

export function forEachLargeSimpleInflatedZipEntryChunk(
  zip: XlsxZipEntries,
  path: string,
  onChunk: (chunk: Uint8Array) => void,
  options: {
    readonly chunkSize?: number
    readonly forceStreamingInflate?: boolean
    readonly garbageCollectionIntervalBytes?: number
    readonly collectGarbage?: () => void
  } = {},
): boolean {
  const collectAfterBytes = Math.max(0, Math.trunc(options.garbageCollectionIntervalBytes ?? defaultGarbageCollectionIntervalBytes))
  const collectGarbage = options.collectGarbage ?? collectLargeSimpleImportGarbage
  let bytesSinceCollection = 0
  return forEachInflatedXlsxZipEntryChunk(
    zip,
    path,
    (chunk) => {
      onChunk(chunk)
      if (collectAfterBytes === 0) {
        return
      }
      bytesSinceCollection += chunk.byteLength
      if (bytesSinceCollection >= collectAfterBytes) {
        bytesSinceCollection = 0
        collectGarbage()
      }
    },
    {
      forceStreamingInflate: options.forceStreamingInflate ?? true,
      ...(options.chunkSize === undefined ? {} : { chunkSize: options.chunkSize }),
    },
  )
}
