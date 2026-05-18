import { parseRangeAddress } from './addressing.js'
import { quoteSheetNameIfNeeded } from './translation-reference-utils.js'

export interface ParsedQualifiedRangeReference {
  readonly sheetName?: string
  readonly sheetEndName?: string
  readonly refKind: 'cells' | 'rows' | 'cols'
  readonly start: string
  readonly end: string
}

export function parseRawQualifiedRangeReference(raw: string): ParsedQualifiedRangeReference {
  const bang = raw.lastIndexOf('!')
  if (bang === -1) {
    const parsed = parseRangeAddress(raw)
    const rawRange = splitLocalRangeReference(raw)
    return {
      refKind: parsed.kind,
      start: rawRange.start,
      end: rawRange.end,
    }
  }

  const qualifier = raw.slice(0, bang).trim()
  const localRangeText = raw.slice(bang + 1).trim()
  const parsedLocalRange = parseRangeAddress(localRangeText)
  const rawRange = splitLocalRangeReference(localRangeText)
  const sheetRange = splitSheetRangeQualifier(qualifier)

  if (sheetRange) {
    return {
      sheetName: unquoteSheetQualifierPart(sheetRange.start),
      sheetEndName: unquoteSheetQualifierPart(sheetRange.end),
      refKind: parsedLocalRange.kind,
      start: rawRange.start,
      end: rawRange.end,
    }
  }

  return {
    sheetName: unquoteSheetQualifierPart(qualifier),
    refKind: parsedLocalRange.kind,
    start: rawRange.start,
    end: rawRange.end,
  }
}

export function formatQualifiedRangeReference(
  sheetName: string | undefined,
  sheetEndName: string | undefined,
  start: string,
  end: string,
): string {
  const prefix =
    sheetName && sheetEndName
      ? `${quoteSheetNameIfNeeded(sheetName)}:${quoteSheetNameIfNeeded(sheetEndName)}!`
      : sheetName
        ? `${quoteSheetNameIfNeeded(sheetName)}!`
        : ''
  return `${prefix}${start}:${end}`
}

function splitLocalRangeReference(raw: string): { readonly start: string; readonly end: string } {
  const separator = raw.indexOf(':')
  if (separator <= 0 || separator >= raw.length - 1) {
    throw new Error(`Invalid range address: ${raw}`)
  }
  return {
    start: raw.slice(0, separator).trim(),
    end: raw.slice(separator + 1).trim(),
  }
}

function splitSheetRangeQualifier(qualifier: string): { readonly start: string; readonly end: string } | undefined {
  let quoted = false
  for (let index = 0; index < qualifier.length; index += 1) {
    const char = qualifier[index]!
    if (char === "'") {
      if (quoted && qualifier[index + 1] === "'") {
        index += 1
        continue
      }
      quoted = !quoted
      continue
    }
    if (char === ':' && !quoted) {
      return {
        start: qualifier.slice(0, index),
        end: qualifier.slice(index + 1),
      }
    }
  }
  return undefined
}

function unquoteSheetQualifierPart(part: string): string {
  const trimmed = part.trim()
  return trimmed.startsWith("'") && trimmed.endsWith("'") ? trimmed.slice(1, -1).replace(/''/g, "'") : trimmed
}
