import { getZipText, readXlsxZipEntries, setZipText } from './zip-reader.js'
import { zipSourcePreservingEntries } from './source-preserving-zip.js'

export interface XlsxTextPartPatch {
  readonly path: string
  patchText(text: string): string
}

export interface XlsxWorksheetCellXmlReplacement {
  readonly path: string
  readonly address: string
  readonly replacement: string
  readonly missingMessage?: string
}

export function patchXlsxTextParts(bytes: Uint8Array, patches: readonly XlsxTextPartPatch[]): Uint8Array {
  if (patches.length === 0) {
    return bytes
  }
  const zip = readXlsxZipEntries(bytes)
  let changed = false
  for (const patch of patches) {
    const text = getZipText(zip, patch.path)
    if (text === null) {
      continue
    }
    const nextText = patch.patchText(text)
    if (nextText === text) {
      continue
    }
    setZipText(zip, patch.path, nextText)
    changed = true
  }
  return changed ? zipSourcePreservingEntries(zip) : bytes
}

export function replaceXlsxWorksheetCellXml(bytes: Uint8Array, replacement: XlsxWorksheetCellXmlReplacement): Uint8Array {
  let replaced = false
  const output = patchXlsxTextParts(bytes, [
    {
      path: replacement.path,
      patchText: (text) => {
        const nextText = text.replace(
          new RegExp(`<c\\b[^>]*\\br="${escapeRegExp(replacement.address)}"[^>]*>[\\s\\S]*?<\\/c>`, 'u'),
          () => {
            replaced = true
            return replacement.replacement
          },
        )
        return nextText
      },
    },
  ])
  if (!replaced) {
    throw new Error(replacement.missingMessage ?? `XLSX worksheet is missing ${replacement.path} ${replacement.address}`)
  }
  return output
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
