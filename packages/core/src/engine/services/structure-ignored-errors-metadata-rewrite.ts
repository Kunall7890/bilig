import { MAX_COLS, MAX_ROWS, type CellRangeRef, type WorkbookIgnoredErrorsSnapshot } from '@bilig/protocol'
import {
  columnToIndex,
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'

const ignoredErrorElementPattern = /<((?:[A-Za-z_][\w.-]*:)?ignoredError)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/\1>)/gu
const ignoredErrorsContainerPattern = /<(?:[A-Za-z_][\w.-]*:)?ignoredErrors\b[\s\S]*<\/(?:[A-Za-z_][\w.-]*:)?ignoredErrors>/u
const sqrefAttributePattern = /\bsqref=(["'])([^"']*)\1/u
const cellReferencePattern = /^\$?([A-Z]+)\$?([1-9]\d*)$/iu

export function rewriteIgnoredErrorsForStructuralTransform(
  ignoredErrors: WorkbookIgnoredErrorsSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookIgnoredErrorsSnapshot | undefined {
  if (!ignoredErrors) {
    return undefined
  }
  const xml = ignoredErrors.xml.replace(
    ignoredErrorElementPattern,
    (_match, tagName: string, attributes: string, body: string | undefined) => {
      const nextAttributes = rewriteIgnoredErrorAttributes(attributes, transform)
      if (nextAttributes === null) {
        return ''
      }
      return body === undefined ? `<${tagName}${nextAttributes}/>` : `<${tagName}${nextAttributes}>${body}</${tagName}>`
    },
  )
  if (!ignoredErrorsContainerPattern.test(xml) || !/<(?:[A-Za-z_][\w.-]*:)?ignoredError\b/u.test(xml)) {
    return undefined
  }
  return { xml }
}

function rewriteIgnoredErrorAttributes(attributes: string, transform: StructuralAxisTransform): string | null {
  const match = sqrefAttributePattern.exec(attributes)
  if (!match) {
    return attributes
  }
  const quote = match[1] ?? '"'
  const refs = (match[2] ?? '')
    .trim()
    .split(/\s+/u)
    .flatMap((token) => rewriteSqrefToken(token, transform))
  if (refs.length === 0) {
    return null
  }
  return attributes.replace(match[0], `sqref=${quote}${refs.join(' ')}${quote}`)
}

function rewriteSqrefToken(token: string, transform: StructuralAxisTransform): string[] {
  if (token.length === 0) {
    return []
  }
  const rangeSeparator = token.indexOf(':')
  if (rangeSeparator >= 0) {
    const startAddress = token.slice(0, rangeSeparator)
    const endAddress = token.slice(rangeSeparator + 1)
    if (!isCellReference(startAddress) || !isCellReference(endAddress)) {
      return [token]
    }
    const rewritten = rewriteRangeForStructuralTransform(startAddress, endAddress, transform)
    if (!rewritten) {
      return []
    }
    const clipped = clipRangeToSheetGrid(rewritten.startAddress, rewritten.endAddress)
    return clipped ? [`${clipped.startAddress}:${clipped.endAddress}`] : []
  }
  if (!isCellReference(token)) {
    return [token]
  }
  const rewritten = rewriteAddressForStructuralTransform(token, transform)
  if (!rewritten) {
    return []
  }
  const clipped = clipAddressToSheetGrid(rewritten)
  return clipped ? [clipped] : []
}

function isCellReference(address: string): boolean {
  return cellReferencePattern.test(address)
}

function clipRangeToSheetGrid(startAddress: string, endAddress: string): CellRangeRef | undefined {
  const start = parseCellAddress(startAddress)
  const end = parseCellAddress(endAddress)
  if (!start || !end) {
    return undefined
  }
  const startRow = Math.min(start[0], end[0])
  const endRow = Math.min(MAX_ROWS - 1, Math.max(start[0], end[0]))
  const startCol = Math.min(start[1], end[1])
  const endCol = Math.min(MAX_COLS - 1, Math.max(start[1], end[1]))
  if (startRow > endRow || startCol > endCol) {
    return undefined
  }
  return {
    sheetName: '',
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  }
}

function clipAddressToSheetGrid(address: string): string | undefined {
  const parsed = parseCellAddress(address)
  if (!parsed) {
    return undefined
  }
  if (parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = cellReferencePattern.exec(address)
  if (!match) {
    return undefined
  }
  return [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())]
}
