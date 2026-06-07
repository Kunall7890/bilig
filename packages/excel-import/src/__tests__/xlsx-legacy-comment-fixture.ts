import { writeSimpleXlsxWorkbook, type SimpleXlsxCell } from '@bilig/xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const commentsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const commentsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'

export interface LegacyCommentFixtureComment {
  readonly ref: string
  readonly author: string
  readonly textXml: string
}

export interface LegacyCommentFixtureShape {
  readonly row: number
  readonly column: number
  readonly anchor: string
  readonly fillColor: string
  readonly marginLeft: string
  readonly marginTop: string
  readonly width: string
  readonly height: string
  readonly visible: boolean
}

export interface LegacyCommentWorkbookFixtureInput {
  readonly sheetName: string
  readonly cells: readonly SimpleXlsxCell[]
  readonly comments: readonly LegacyCommentFixtureComment[]
  readonly shapes: readonly LegacyCommentFixtureShape[]
}

export function buildLegacyCommentWorkbookBytes(input: LegacyCommentWorkbookFixtureInput): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [{ name: input.sheetName, cells: input.cells }],
    }),
  )

  zip['xl/worksheets/sheet1.xml'] = strToU8(addLegacyDrawing(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())))
  zip['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Relationships xmlns="${relationshipNamespace}">`,
      `<Relationship Id="rIdLegacyCommentVml1" Type="${vmlDrawingRelationshipType}" Target="../drawings/vmlDrawing1.vml"/>`,
      `<Relationship Id="rIdLegacyComments1" Type="${commentsRelationshipType}" Target="../comments1.xml"/>`,
      '</Relationships>',
    ].join(''),
  )
  zip['xl/comments1.xml'] = strToU8(commentsXml(input.comments))
  zip['xl/drawings/vmlDrawing1.vml'] = strToU8(vmlDrawingXml(input.shapes))
  zip['[Content_Types].xml'] = strToU8(
    addContentTypeOverride(
      addContentTypeDefault(strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array()), 'vml', vmlDrawingContentType),
      '/xl/comments1.xml',
      commentsContentType,
    ),
  )
  return zipSync(zip)
}

function addLegacyDrawing(sheetXml: string): string {
  return sheetXml.replace('</worksheet>', '<legacyDrawing r:id="rIdLegacyCommentVml1"/></worksheet>')
}

function commentsXml(comments: readonly LegacyCommentFixtureComment[]): string {
  const authors = Array.from(new Set(comments.map((comment) => comment.author)))
  const authorIdByName = new Map(authors.map((author, index) => [author, index]))
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<authors>',
    ...authors.map((author) => `<author>${escapeXmlText(author)}</author>`),
    '</authors>',
    '<commentList>',
    ...comments.map((comment) => {
      const authorId = authorIdByName.get(comment.author) ?? 0
      return `<comment ref="${escapeXmlAttribute(comment.ref)}" authorId="${String(authorId)}"><text>${comment.textXml}</text></comment>`
    }),
    '</commentList>',
    '</comments>',
  ].join('')
}

function vmlDrawingXml(shapes: readonly LegacyCommentFixtureShape[]): string {
  return [
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
    '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
    '</v:shapetype>',
    ...shapes.map((shape, index) => noteShapeXml(shape, index)),
    '</xml>',
  ].join('')
}

function noteShapeXml(shape: LegacyCommentFixtureShape, index: number): string {
  const style = [
    'position:absolute',
    `margin-left:${shape.marginLeft}`,
    `margin-top:${shape.marginTop}`,
    `width:${shape.width}`,
    `height:${shape.height}`,
    'z-index:1',
    `visibility:${shape.visible ? 'visible' : 'hidden'}`,
  ].join(';')
  return [
    `<v:shape id="_x0000_s${String(1025 + index)}" type="#_x0000_t202" style="${escapeXmlAttribute(style)}" fillcolor="${escapeXmlAttribute(
      shape.fillColor,
    )}" o:insetmode="auto">`,
    '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>',
    '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>',
    '<x:ClientData ObjectType="Note">',
    shape.visible ? '<x:Visible/>' : '',
    `<x:Anchor>${escapeXmlText(shape.anchor)}</x:Anchor>`,
    `<x:Row>${String(shape.row)}</x:Row>`,
    `<x:Column>${String(shape.column)}</x:Column>`,
    '</x:ClientData>',
    '</v:shape>',
  ].join('')
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default\\b[^>]*\\bExtension=(["'])${escapeRegExp(extension)}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXmlAttribute(extension)}" ContentType="${escapeXmlAttribute(contentType)}"/></Types>`,
  )
}

function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (new RegExp(`<Override\\b[^>]*\\bPartName=(["'])${escapeRegExp(partName)}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${escapeXmlAttribute(partName)}" ContentType="${escapeXmlAttribute(contentType)}"/></Types>`,
  )
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function escapeXmlText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}
