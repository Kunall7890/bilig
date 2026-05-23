import { columnToIndex, formatAddress, rewriteAddressForStructuralTransform, type StructuralAxisTransform } from '@bilig/formula'
import { MAX_COLS, MAX_ROWS, type WorkbookCommentThreadSnapshot } from '@bilig/protocol'
import type { WorkbookStore } from '../../workbook-store.js'

const CELL_REF_RE = /^\$?([A-Z]+)\$?([1-9]\d*)$/i

export function rewriteLegacyCommentVmlForStructuralTransform(args: {
  readonly workbook: WorkbookStore
  readonly sheetName: string
  readonly transform: StructuralAxisTransform
}): void {
  const legacyCommentVml = args.workbook.getSheetLegacyCommentVml(args.sheetName)
  if (!legacyCommentVml) {
    return
  }

  const nextCommentsXml = legacyCommentVml.commentsXml
    ? rewriteCommentXmlForStructuralTransform(legacyCommentVml.commentsXml, args.transform)
    : undefined
  const nextVmlXml = rewriteVmlXmlForStructuralTransform(legacyCommentVml.vmlXml, args.transform)
  const nextCommentSignature = legacyCommentThreadSignature(args.workbook.listCommentThreads(args.sheetName))

  if (
    nextCommentsXml === legacyCommentVml.commentsXml &&
    nextVmlXml === legacyCommentVml.vmlXml &&
    nextCommentSignature === legacyCommentVml.commentSignature
  ) {
    return
  }

  args.workbook.setSheetLegacyCommentVml(args.sheetName, {
    relationshipTarget: legacyCommentVml.relationshipTarget,
    vmlXml: nextVmlXml,
    ...(legacyCommentVml.commentsRelationshipTarget !== undefined
      ? { commentsRelationshipTarget: legacyCommentVml.commentsRelationshipTarget }
      : {}),
    ...(nextCommentsXml !== undefined ? { commentsXml: nextCommentsXml } : {}),
    commentSignature: nextCommentSignature,
  })
}

function rewriteCommentXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  return xml.replace(/<comment\b([^>]*?)(\/>|>[\s\S]*?<\/comment>)/gu, (source: string, attributes: string) => {
    const ref = readXmlAttribute(attributes, 'ref')
    if (!ref) {
      return source
    }
    const nextRef = rewriteCellRefForStructuralTransform(ref, transform)
    if (!nextRef) {
      return ''
    }
    return source.replace(/\bref=(["'])([\s\S]*?)\1/u, (_attribute: string, quote: string) => `ref=${quote}${nextRef}${quote}`)
  })
}

function rewriteVmlXmlForStructuralTransform(xml: string, transform: StructuralAxisTransform): string {
  let nextXml = xml
  if (transform.axis === 'row') {
    nextXml = nextXml.replace(/<x:Row>(\d+)<\/x:Row>/gu, (source: string, rowText: string) => {
      const nextRow = rewriteZeroBasedRowForStructuralTransform(Number(rowText), transform)
      return nextRow === undefined ? source : `<x:Row>${String(nextRow)}</x:Row>`
    })
  } else {
    nextXml = nextXml.replace(/<x:Column>(\d+)<\/x:Column>/gu, (source: string, columnText: string) => {
      const nextColumn = rewriteZeroBasedColumnForStructuralTransform(Number(columnText), transform)
      return nextColumn === undefined ? source : `<x:Column>${String(nextColumn)}</x:Column>`
    })
  }

  return nextXml.replace(/<x:Anchor>([\s\S]*?)<\/x:Anchor>/gu, (source: string, body: string) => {
    const values = body
      .split(',')
      .map((value) => Number(value.trim()))
      .filter((value) => Number.isInteger(value))
    if (values.length !== 8) {
      return source
    }

    const firstAxisIndex = transform.axis === 'row' ? 2 : 0
    const secondAxisIndex = transform.axis === 'row' ? 6 : 4
    const first = rewriteZeroBasedIndexForStructuralTransform(values[firstAxisIndex]!, transform)
    const second = rewriteZeroBasedIndexForStructuralTransform(values[secondAxisIndex]!, transform)
    if (first === undefined || second === undefined) {
      return source
    }

    values[firstAxisIndex] = first
    values[secondAxisIndex] = second
    return `<x:Anchor>${values.join(', ')}</x:Anchor>`
  })
}

function rewriteZeroBasedIndexForStructuralTransform(index: number, transform: StructuralAxisTransform): number | undefined {
  return transform.axis === 'row'
    ? rewriteZeroBasedRowForStructuralTransform(index, transform)
    : rewriteZeroBasedColumnForStructuralTransform(index, transform)
}

function rewriteZeroBasedRowForStructuralTransform(row: number, transform: StructuralAxisTransform): number | undefined {
  const rewritten = rewriteAddressForStructuralTransform(`A${String(row + 1)}`, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseCellAddress(rewritten)
  return parsed && parsed[0] < MAX_ROWS ? parsed[0] : undefined
}

function rewriteZeroBasedColumnForStructuralTransform(column: number, transform: StructuralAxisTransform): number | undefined {
  const rewritten = rewriteAddressForStructuralTransform(formatAddress(0, column), transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseCellAddress(rewritten)
  return parsed && parsed[1] < MAX_COLS ? parsed[1] : undefined
}

function rewriteCellRefForStructuralTransform(ref: string, transform: StructuralAxisTransform): string | undefined {
  const rewritten = rewriteAddressForStructuralTransform(ref, transform)
  if (!rewritten) {
    return undefined
  }
  const parsed = parseCellAddress(rewritten)
  if (!parsed || parsed[0] >= MAX_ROWS || parsed[1] >= MAX_COLS) {
    return undefined
  }
  return formatAddress(parsed[0], parsed[1])
}

function parseCellAddress(address: string): [number, number] | undefined {
  const match = CELL_REF_RE.exec(address)
  return match ? [+match[2]! - 1, columnToIndex(match[1]!.toUpperCase())] : undefined
}

function normalizeCommentAddress(address: string): string {
  const parsed = parseCellAddress(address)
  return parsed ? formatAddress(parsed[0], parsed[1]) : address.trim().toUpperCase()
}

function legacyCommentThreadSignature(commentThreads: readonly WorkbookCommentThreadSnapshot[]): string {
  const normalized = commentThreads
    .map((thread) => ({
      sheetName: thread.sheetName,
      address: normalizeCommentAddress(thread.address),
      comments: thread.comments.map((comment) => ({
        body: comment.body,
        authorDisplayName: comment.authorDisplayName ?? '',
      })),
    }))
    .toSorted((left, right) => `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`))
  return JSON.stringify(normalized)
}

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}
