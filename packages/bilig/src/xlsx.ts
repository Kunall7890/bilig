export * from '@bilig/headless/xlsx'

import { importXlsx as importHeadlessXlsx, type XlsxImportOptions } from '@bilig/headless/xlsx'

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}) {
  return importHeadlessXlsx(bytes, fileName, {
    preferNativeSimpleImport: true,
    ...options,
  })
}
