import { strFromU8, unzipSync, type Unzipped } from 'fflate'

export type XlsxZipEntries = Unzipped
export type XlsxZipSource = Uint8Array | XlsxZipEntries

export function readXlsxZipEntries(source: XlsxZipSource): XlsxZipEntries {
  return source instanceof Uint8Array ? unzipSync(source) : source
}

export function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

export function getZipText(zip: XlsxZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}
