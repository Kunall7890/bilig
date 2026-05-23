import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const threadedCommentRelationshipType = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment'
const personRelationshipType = 'http://schemas.microsoft.com/office/2017/10/relationships/person'
const commentsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const threadedCommentContentType = 'application/vnd.ms-excel.threadedcomments+xml'
const personContentType = 'application/vnd.ms-excel.person+xml'
const commentsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const excelAuthoredThreadedCommentFixtureUrl = new URL(
  '../../fixtures/xlsx-corpus/macos-excel-threaded-comments-source.xlsx',
  import.meta.url,
)

describe('macOS Desktop Excel threaded comment structural oracle', () => {
  it('preserves threaded comment package artifacts and rewrites refs after headless row inserts', () => {
    const imported = importXlsx(threadedCommentWorkbookBytes(), 'threaded-comment-source.xlsx')
    expect(imported.snapshot.workbook.metadata?.threadedCommentArtifacts?.parts).toHaveLength(2)
    expect(imported.snapshot.sheets[0]?.metadata?.threadedCommentArtifacts?.relationships).toHaveLength(1)
    expect(imported.snapshot.sheets[0]?.metadata?.legacyCommentVml).toBeDefined()

    const workpaper = WorkPaper.buildFromSnapshot(imported.snapshot)
    try {
      const review = workpaper.getSheetId('Review')
      if (review === undefined) {
        throw new Error('Expected Review sheet to be available')
      }
      workpaper.addRows(review, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      expect(readThreadedCommentRefs(exported)).toEqual(['A3'])
      expect(readLegacyCommentRefs(exported)).toEqual(['A3'])
      expect(readPersonNames(exported)).toEqual(['Finance Reviewer'])
      expect(readThreadedCommentRelationshipCount(exported)).toBe(1)
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel threaded comment refs after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-threaded-comments-oracle-')
      try {
        const sourceBytes = new Uint8Array(readFileSync(excelAuthoredThreadedCommentFixtureUrl))
        expect(readThreadedCommentRefs(sourceBytes)).toEqual(['B1'])
        expect(readLegacyCommentRefs(sourceBytes)).toEqual(['B1'])

        const sourcePath = join(tempDir, 'threaded-comment-source.xlsx')
        writeFileSync(sourcePath, sourceBytes)

        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: sourcePath,
          worksheetName: 'Country equity risk premiums',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['B1', 'B2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'number', value: 0.0423 },
        ])
        expect(readThreadedCommentRefs(new Uint8Array(readFileSync(sourcePath)))).toEqual(['B2'])
        expect(readLegacyCommentRefs(new Uint8Array(readFileSync(sourcePath)))).toEqual(['B2'])

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(sourceBytes, 'excel-authored-threaded-comments-source.xlsx').snapshot)
        try {
          const countryErp = workpaper.getSheetId('Country equity risk premiums')
          if (countryErp === undefined) {
            throw new Error('Expected Country equity risk premiums sheet to be available')
          }
          workpaper.addRows(countryErp, 0, 1)

          const headlessPath = join(tempDir, 'headless-threaded-comment-oracle.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          expect(readThreadedCommentRefs(new Uint8Array(readFileSync(headlessPath)))).toEqual(['B2'])
          expect(readLegacyCommentRefs(new Uint8Array(readFileSync(headlessPath)))).toEqual(['B2'])
          expect(readPersonNames(new Uint8Array(readFileSync(headlessPath)))).toEqual(readPersonNames(sourceBytes))
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function threadedCommentWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Assumption'], ['Needs review']]), 'Review')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))

  zip['xl/threadedComments/threadedComment1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<threadedComment ref="A2" dT="2026-05-01T10:00:00Z" personId="{11111111-1111-1111-1111-111111111111}" id="{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}"><text><x:r><x:t>Check assumption before filing</x:t></x:r></text></threadedComment>',
      '</ThreadedComments>',
    ].join(''),
  )
  zip['xl/persons/person1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/person">',
      '<person displayName="Finance Reviewer" id="{11111111-1111-1111-1111-111111111111}" userId="finance@example.com" providerId="None"/>',
      '</personList>',
    ].join(''),
  )
  zip['xl/comments1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">',
      '<authors><author>tc={aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}</author></authors>',
      '<commentList>',
      '<comment ref="A2" authorId="0" shapeId="0" xr:uid="{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}"><text><t>[Threaded comment]\n\nComment:\n    Check assumption before filing</t></text></comment>',
      '</commentList>',
      '</comments>',
    ].join(''),
  )
  zip['xl/drawings/vmlDrawing1.vml'] = strToU8(
    [
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
      '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>',
      '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:60pt;margin-top:30pt;width:144pt;height:72pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">',
      '<v:fill color2="#ffffe1"/><v:shadow on="t" obscured="t"/><v:path o:connecttype="none"/><v:textbox><div style="text-align:left"></div></v:textbox>',
      '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>0, 15, 1, 2, 2, 15, 4, 2</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData>',
      '</v:shape>',
      '</xml>',
    ].join(''),
  )
  zip['xl/worksheets/sheet1.xml'] = addLegacyDrawingToWorksheet(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()), 'rIdVml1')

  upsertRelationship(zip, 'xl/worksheets/_rels/sheet1.xml.rels', {
    id: 'rIdThreaded1',
    type: threadedCommentRelationshipType,
    target: '../threadedComments/threadedComment1.xml',
  })
  upsertRelationship(zip, 'xl/worksheets/_rels/sheet1.xml.rels', {
    id: 'rIdComments1',
    type: commentsRelationshipType,
    target: '../comments1.xml',
  })
  upsertRelationship(zip, 'xl/worksheets/_rels/sheet1.xml.rels', {
    id: 'rIdVml1',
    type: vmlDrawingRelationshipType,
    target: '../drawings/vmlDrawing1.vml',
  })
  upsertRelationship(zip, 'xl/_rels/workbook.xml.rels', {
    id: 'rIdPerson1',
    type: personRelationshipType,
    target: 'persons/person1.xml',
  })
  addContentTypeOverride(zip, '/xl/threadedComments/threadedComment1.xml', threadedCommentContentType)
  addContentTypeOverride(zip, '/xl/persons/person1.xml', personContentType)
  addContentTypeOverride(zip, '/xl/comments1.xml', commentsContentType)
  addContentTypeDefault(zip, 'vml', vmlDrawingContentType)
  return zipSync(zip)
}

function addLegacyDrawingToWorksheet(worksheetXml: string, relationshipId: string): Uint8Array {
  if (worksheetXml.includes('<legacyDrawing')) {
    return strToU8(worksheetXml)
  }
  return strToU8(worksheetXml.replace('</worksheet>', `<legacyDrawing r:id="${relationshipId}"/></worksheet>`))
}

function upsertRelationship(
  zip: Record<string, Uint8Array>,
  relsPath: string,
  relationship: {
    readonly id: string
    readonly type: string
    readonly target: string
  },
): void {
  const relationshipXml = `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`
  const currentXml = strFromU8(zip[relsPath] ?? new Uint8Array())
  zip[relsPath] = strToU8(
    currentXml.includes('</Relationships>')
      ? currentXml.replace('</Relationships>', `${relationshipXml}</Relationships>`)
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">${relationshipXml}</Relationships>`,
  )
}

function addContentTypeOverride(zip: Record<string, Uint8Array>, partName: string, contentType: string): void {
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.includes(`PartName="${partName}"`)
      ? contentTypesXml
      : contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`),
  )
}

function addContentTypeDefault(zip: Record<string, Uint8Array>, extension: string, contentType: string): void {
  const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())
  zip['[Content_Types].xml'] = strToU8(
    contentTypesXml.includes(`Extension="${extension}"`)
      ? contentTypesXml
      : contentTypesXml.replace('</Types>', `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`),
  )
}

function readThreadedCommentRefs(bytes: Uint8Array): string[] {
  const zip = unzipSync(bytes)
  return Object.entries(zip)
    .filter(([path]) => /^xl\/threadedComments\/threadedComment[^/]*\.xml$/u.test(path))
    .flatMap(([_path, part]) =>
      [...strFromU8(part).matchAll(/<threadedComment\b[^>]*\bref=(["'])(.*?)\1/gu)].map((match) => match[2] ?? ''),
    )
    .toSorted()
}

function readLegacyCommentRefs(bytes: Uint8Array): string[] {
  const zip = unzipSync(bytes)
  return Object.entries(zip)
    .filter(([path]) => /^xl\/comments[^/]*\.xml$/u.test(path))
    .flatMap(([_path, part]) => [...strFromU8(part).matchAll(/<comment\b[^>]*\bref=(["'])(.*?)\1/gu)].map((match) => match[2] ?? ''))
    .toSorted()
}

function readPersonNames(bytes: Uint8Array): string[] {
  const zip = unzipSync(bytes)
  return Object.entries(zip)
    .filter(([path]) => /^xl\/persons\/person[^/]*\.xml$/u.test(path))
    .flatMap(([_path, part]) => [...strFromU8(part).matchAll(/\bdisplayName=(["'])(.*?)\1/gu)].map((match) => match[2] ?? ''))
    .toSorted()
}

function readThreadedCommentRelationshipCount(bytes: Uint8Array): number {
  const zip = unzipSync(bytes)
  return Object.entries(zip)
    .filter(([path]) => /^xl\/worksheets\/_rels\/sheet[^/]*\.xml\.rels$/u.test(path))
    .reduce((count, [_path, rels]) => count + (strFromU8(rels).match(new RegExp(threadedCommentRelationshipType, 'gu'))?.length ?? 0), 0)
}
