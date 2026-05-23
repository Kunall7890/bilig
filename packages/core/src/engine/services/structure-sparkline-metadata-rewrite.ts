import {
  rewriteAddressForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
} from '@bilig/formula'
import type { WorkbookSparklinesSnapshot } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'

function decodeXmlText(value: string): string {
  return value
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, '<')
    .replace(/&gt;/gu, '>')
    .replace(/&amp;/gu, '&')
}

function escapeXmlText(value: string): string {
  return value.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;')
}

function rewriteSparklineReference(value: string, transform: StructuralAxisTransform): string | undefined {
  const [start, end, extra] = value.split(':')
  if (!start || extra !== undefined) {
    return undefined
  }
  try {
    if (!end) {
      return rewriteAddressForStructuralTransform(start, transform)
    }
    const rewritten = rewriteRangeForStructuralTransform(start, end, transform)
    return rewritten ? `${rewritten.startAddress}:${rewritten.endAddress}` : undefined
  } catch {
    return undefined
  }
}

function rewriteSparklineSqref(value: string, transform: StructuralAxisTransform): string | undefined {
  const refs = value.trim().split(/\s+/u).filter(Boolean)
  if (refs.length === 0) {
    return undefined
  }
  const rewritten = refs.flatMap((ref) => {
    const nextRef = rewriteSparklineReference(ref, transform)
    return nextRef ? [nextRef] : []
  })
  return rewritten.length > 0 ? rewritten.join(' ') : undefined
}

function rewriteSparklineElementText(
  xml: string,
  elementName: 'f' | 'sqref',
  rewrite: (value: string) => string | undefined,
): { readonly xml: string; readonly failed: boolean } {
  let failed = false
  const pattern = new RegExp(
    `(<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>)([\\s\\S]*?)(</(?:[A-Za-z_][\\w.-]*:)?${elementName}>)`,
    'gu',
  )
  const nextXml = xml.replace(pattern, (match: string, open: string, text: string, close: string) => {
    const nextText = rewrite(decodeXmlText(text))
    if (nextText === undefined) {
      failed = true
      return match
    }
    return `${open}${escapeXmlText(nextText)}${close}`
  })
  return { xml: nextXml, failed }
}

export function rewriteSparklinesForStructuralTransform(
  sheetName: string,
  sparklines: WorkbookSparklinesSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookSparklinesSnapshot | undefined {
  if (!sparklines) {
    return undefined
  }
  const withFormulas = rewriteSparklineElementText(sparklines.xml, 'f', (formula) => {
    try {
      return rewriteFormulaForStructuralTransform(formula, sheetName, sheetName, transform)
    } catch {
      return undefined
    }
  })
  if (withFormulas.failed) {
    return undefined
  }
  const withSqrefs = rewriteSparklineElementText(withFormulas.xml, 'sqref', (sqref) => rewriteSparklineSqref(sqref, transform))
  return withSqrefs.failed ? undefined : { xml: withSqrefs.xml }
}
