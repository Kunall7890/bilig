import { MAX_COLS, MAX_ROWS, type WorkbookSheetPrintPageSetupSnapshot } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'

const breakElementPattern = /<brk\b([^>]*?)\/>/gu
const idAttributePattern = /\bid=(["'])(\d+)\1/u

export function rewritePrintPageSetupForStructuralTransform(
  printPageSetup: WorkbookSheetPrintPageSetupSnapshot | undefined,
  transform: StructuralAxisTransform,
): WorkbookSheetPrintPageSetupSnapshot | undefined {
  if (!printPageSetup) {
    return undefined
  }
  const rewritten: WorkbookSheetPrintPageSetupSnapshot = { ...printPageSetup }
  if (transform.axis === 'row' && printPageSetup.rowBreaksXml) {
    const rowBreaksXml = rewriteBreaksXml(printPageSetup.rowBreaksXml, 'row', transform)
    if (rowBreaksXml) {
      rewritten.rowBreaksXml = rowBreaksXml
    } else {
      delete rewritten.rowBreaksXml
    }
  }
  if (transform.axis === 'column' && printPageSetup.colBreaksXml) {
    const colBreaksXml = rewriteBreaksXml(printPageSetup.colBreaksXml, 'column', transform)
    if (colBreaksXml) {
      rewritten.colBreaksXml = colBreaksXml
    } else {
      delete rewritten.colBreaksXml
    }
  }
  return Object.keys(rewritten).length > 0 ? rewritten : undefined
}

function rewriteBreaksXml(xml: string, axis: 'row' | 'column', transform: StructuralAxisTransform): string | undefined {
  let breakCount = 0
  let manualBreakCount = 0
  const rewritten = xml.replace(breakElementPattern, (_match, attributes: string) => {
    const nextAttributes = rewriteBreakAttributes(attributes, axis, transform)
    if (!nextAttributes) {
      return ''
    }
    breakCount += 1
    if (isManualBreak(nextAttributes)) {
      manualBreakCount += 1
    }
    return `<brk${nextAttributes}/>`
  })
  if (breakCount === 0) {
    return undefined
  }
  return updateBreakCounts(rewritten, breakCount, manualBreakCount)
}

function rewriteBreakAttributes(attributes: string, axis: 'row' | 'column', transform: StructuralAxisTransform): string | undefined {
  const match = idAttributePattern.exec(attributes)
  if (!match) {
    return attributes
  }
  const nextId = rewriteBreakId(Number(match[2]), axis, transform)
  if (nextId === undefined) {
    return undefined
  }
  return attributes.replace(match[0], `id=${match[1] ?? '"'}${String(nextId)}${match[1] ?? '"'}`)
}

function rewriteBreakId(id: number, axis: 'row' | 'column', transform: StructuralAxisTransform): number | undefined {
  if (!Number.isInteger(id) || id < 0 || axis !== transform.axis) {
    return undefined
  }
  let nextId: number | undefined
  switch (transform.kind) {
    case 'insert':
      nextId = id >= transform.start ? id + transform.count : id
      break
    case 'delete':
      if (id < transform.start) {
        nextId = id
      } else if (id >= transform.start + transform.count) {
        nextId = id - transform.count
      }
      break
    case 'move':
      nextId = rewriteMovedBreakId(id, transform)
      break
  }
  if (nextId === undefined || nextId >= (axis === 'row' ? MAX_ROWS : MAX_COLS)) {
    return undefined
  }
  return nextId
}

function rewriteMovedBreakId(id: number, transform: Extract<StructuralAxisTransform, { readonly kind: 'move' }>): number {
  if (transform.target < transform.start) {
    if (id >= transform.target && id < transform.start) {
      return id + transform.count
    }
  } else if (transform.target > transform.start) {
    if (id >= transform.start + transform.count && id < transform.target + transform.count) {
      return id - transform.count
    }
  }
  if (id >= transform.start && id < transform.start + transform.count) {
    return transform.target + (id - transform.start)
  }
  return id
}

function isManualBreak(attributes: string): boolean {
  return /\bman=(["'])1\1/u.test(attributes)
}

function updateBreakCounts(xml: string, breakCount: number, manualBreakCount: number): string {
  return xml.replace(/<((?:[A-Za-z_][\w.-]*:)?(?:rowBreaks|colBreaks))\b([^>]*)>/u, (_match, tagName: string, attributes: string) => {
    const withCount = setCountAttribute(attributes, 'count', breakCount)
    const nextAttributes = /\bmanualBreakCount=/u.test(withCount)
      ? setCountAttribute(withCount, 'manualBreakCount', manualBreakCount)
      : withCount
    return `<${tagName}${nextAttributes}>`
  })
}

function setCountAttribute(attributes: string, name: 'count' | 'manualBreakCount', value: number): string {
  if (new RegExp(`\\b${name}=`, 'u').test(attributes)) {
    return attributes.replace(new RegExp(`\\b${name}=(["'])\\d+\\1`, 'u'), `${name}="${String(value)}"`)
  }
  return `${attributes} ${name}="${String(value)}"`
}
