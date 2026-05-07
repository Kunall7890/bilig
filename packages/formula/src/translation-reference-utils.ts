const CELL_REF_RE = /^(\$?)([A-Z]+)(\$?)([1-9][0-9]*)$/
const COLUMN_REF_RE = /^(\$?)([A-Z]+)$/
const ROW_REF_RE = /^(\$?)([1-9][0-9]*)$/

export type StructuralAxisKind = 'row' | 'column'

export type StructuralAxisTransform =
  | { kind: 'insert'; axis: StructuralAxisKind; start: number; count: number }
  | { kind: 'delete'; axis: StructuralAxisKind; start: number; count: number }
  | { kind: 'move'; axis: StructuralAxisKind; start: number; count: number; target: number }

export interface ParsedCellReference {
  colAbsolute: boolean
  rowAbsolute: boolean
  col: number
  row: number
}

export interface ParsedAxisReference {
  absolute: boolean
  index: number
}

function assertNever(value: never): never {
  throw new Error(`Unexpected value: ${String(value)}`)
}

export function formatSheetPrefix(sheetName?: string): string {
  if (!sheetName) {
    return ''
  }
  return `${quoteSheetNameIfNeeded(sheetName)}!`
}

export function quoteSheetNameIfNeeded(sheetName: string): string {
  return /^[A-Za-z0-9_.$]+$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

export function columnToIndex(column: string): number {
  let value = 0
  for (const char of column) {
    value = value * 26 + (char.charCodeAt(0) - 64)
  }
  return value - 1
}

export function indexToColumn(index: number): string {
  let current = index + 1
  let output = ''
  while (current > 0) {
    const remainder = (current - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    current = Math.floor((current - 1) / 26)
  }
  return output
}

export function targetsSheet(explicitSheetName: string | undefined, ownerSheetName: string, targetSheetName: string): boolean {
  return (explicitSheetName ?? ownerSheetName) === targetSheetName
}

export function parseCellReferenceParts(ref: string): ParsedCellReference | undefined {
  const match = CELL_REF_RE.exec(ref.toUpperCase())
  if (!match) {
    return undefined
  }
  const [, colAbsolute, columnText, rowAbsolute, rowText] = match
  return {
    colAbsolute: colAbsolute === '$',
    rowAbsolute: rowAbsolute === '$',
    col: columnToIndex(columnText!),
    row: Number.parseInt(rowText!, 10) - 1,
  }
}

export function formatCellReference(parts: ParsedCellReference, row: number, col: number): string {
  return `${parts.colAbsolute ? '$' : ''}${indexToColumn(col)}${parts.rowAbsolute ? '$' : ''}${row + 1}`
}

export function parseAxisReferenceParts(ref: string, kind: StructuralAxisKind): ParsedAxisReference | undefined {
  const match = (kind === 'row' ? ROW_REF_RE : COLUMN_REF_RE).exec(ref.toUpperCase())
  if (!match) {
    return undefined
  }
  return kind === 'row'
    ? {
        absolute: match[1] === '$',
        index: Number.parseInt(match[2]!, 10) - 1,
      }
    : {
        absolute: match[1] === '$',
        index: columnToIndex(match[2]!),
      }
}

export function formatAxisReference(absolute: boolean, index: number, kind: StructuralAxisKind): string {
  const prefix = absolute ? '$' : ''
  return kind === 'row' ? `${prefix}${index + 1}` : `${prefix}${indexToColumn(index)}`
}

export function mapPointIndex(index: number, transform: StructuralAxisTransform): number | undefined {
  switch (transform.kind) {
    case 'insert':
      return index >= transform.start ? index + transform.count : index
    case 'delete':
      if (index < transform.start) {
        return index
      }
      if (index >= transform.start + transform.count) {
        return index - transform.count
      }
      return undefined
    case 'move':
      if (transform.target < transform.start) {
        if (index >= transform.target && index < transform.start) {
          return index + transform.count
        }
      } else if (transform.target > transform.start) {
        if (index >= transform.start + transform.count && index < transform.target + transform.count) {
          return index - transform.count
        }
      }
      if (index >= transform.start && index < transform.start + transform.count) {
        return transform.target + (index - transform.start)
      }
      return index
    default:
      return assertNever(transform)
  }
}

export function mapInterval(start: number, end: number, transform: StructuralAxisTransform): { start: number; end: number } | undefined {
  switch (transform.kind) {
    case 'insert': {
      if (transform.start <= start) {
        return { start: start + transform.count, end: end + transform.count }
      }
      if (transform.start <= end) {
        return { start, end: end + transform.count }
      }
      return { start, end }
    }
    case 'delete': {
      const deleteEnd = transform.start + transform.count - 1
      if (deleteEnd < start) {
        return { start: start - transform.count, end: end - transform.count }
      }
      if (transform.start > end) {
        return { start, end }
      }
      const survivingStart = start < transform.start ? start : deleteEnd + 1
      const survivingEnd = end > deleteEnd ? end : transform.start - 1
      if (survivingStart > survivingEnd) {
        return undefined
      }
      const nextStart = mapPointIndex(survivingStart, transform)
      const nextEnd = mapPointIndex(survivingEnd, transform)
      return nextStart === undefined || nextEnd === undefined ? undefined : { start: nextStart, end: nextEnd }
    }
    case 'move': {
      const segments =
        transform.target < transform.start
          ? [
              { start: 0, end: transform.target - 1, delta: 0 },
              { start: transform.target, end: transform.start - 1, delta: transform.count },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              { start: transform.start + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
          : [
              { start: 0, end: transform.start - 1, delta: 0 },
              {
                start: transform.start,
                end: transform.start + transform.count - 1,
                delta: transform.target - transform.start,
              },
              {
                start: transform.start + transform.count,
                end: transform.target + transform.count - 1,
                delta: -transform.count,
              },
              { start: transform.target + transform.count, end: Number.MAX_SAFE_INTEGER, delta: 0 },
            ]
      let nextStart: number | undefined
      let nextEnd: number | undefined
      segments.forEach((segment) => {
        const overlapStart = Math.max(start, segment.start)
        const overlapEnd = Math.min(end, segment.end)
        if (overlapStart > overlapEnd) {
          return
        }
        const mappedStart = overlapStart + segment.delta
        const mappedEnd = overlapEnd + segment.delta
        nextStart = nextStart === undefined ? mappedStart : Math.min(nextStart, mappedStart)
        nextEnd = nextEnd === undefined ? mappedEnd : Math.max(nextEnd, mappedEnd)
      })
      if (nextStart === undefined || nextEnd === undefined) {
        return undefined
      }
      return { start: nextStart, end: nextEnd }
    }
    default:
      return assertNever(transform)
  }
}
