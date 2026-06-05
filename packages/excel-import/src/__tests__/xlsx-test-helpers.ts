import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

export function readXlsxTestZipText(bytes: Uint8Array, path: string): string {
  return strFromU8(unzipSync(bytes)[path] ?? new Uint8Array())
}

export function patchXlsxTestZipText(bytes: Uint8Array, path: string, patch: (source: string) => string): Uint8Array {
  const zip = unzipSync(bytes)
  zip[path] = strToU8(patch(strFromU8(zip[path] ?? new Uint8Array())))
  return zipSync(zip)
}
