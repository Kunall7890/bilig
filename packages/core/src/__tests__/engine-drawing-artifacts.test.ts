import { describe, expect, it } from 'vitest'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('engine drawing artifact metadata', () => {
  it('round-trips imported workbook and sheet drawing artifacts through engine mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'drawing-artifact-spec' })
    await engine.ready()

    engine.importSnapshot(importedDrawingArtifactSnapshot())
    engine.setCellValue('Sheet1', 'B1', 'headless edit')

    const exported = engine.exportSnapshot()

    expect(exported.workbook.metadata?.drawingArtifacts).toEqual({
      parts: [
        {
          path: 'xl/drawings/drawing1.xml',
          storage: 'base64',
          dataBase64: encodedTextPart(drawingXml),
          byteLength: new TextEncoder().encode(drawingXml).byteLength,
        },
      ],
      contentTypeOverrides: [
        {
          partName: '/xl/drawings/drawing1.xml',
          contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
        },
      ],
    })
    expect(exported.sheets[0]?.metadata?.drawingArtifacts).toEqual({
      relationshipTarget: '../drawings/drawing1.xml',
      preservedChartRelationshipIds: ['rId7'],
    })
  })

  it('structurally rewrites raw DrawingML anchors after row and column inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'drawing-artifact-structural-spec' })
    await engine.ready()

    engine.importSnapshot(importedDrawingArtifactSnapshot())
    engine.insertRows('Sheet1', 0, 1)
    engine.insertColumns('Sheet1', 0, 1)

    const exported = engine.exportSnapshot()
    expect(drawingAxisMarkers(exported, 'row')).toEqual([1, 5])
    expect(drawingAxisMarkers(exported, 'col')).toEqual([1, 3])
  })
})

function importedDrawingArtifactSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Drawing artifact spec',
      metadata: {
        drawingArtifacts: {
          parts: [
            {
              path: 'xl/drawings/drawing1.xml',
              storage: 'base64',
              dataBase64: encodedTextPart(drawingXml),
              byteLength: new TextEncoder().encode(drawingXml).byteLength,
            },
          ],
          contentTypeOverrides: [
            {
              partName: '/xl/drawings/drawing1.xml',
              contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
            },
          ],
        },
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        metadata: {
          drawingArtifacts: {
            relationshipTarget: '../drawings/drawing1.xml',
            preservedChartRelationshipIds: ['rId7'],
          },
        },
        cells: [{ address: 'A1', value: 'logo' }],
      },
    ],
  }
}

function drawingAxisMarkers(snapshot: WorkbookSnapshot, axis: 'row' | 'col'): number[] {
  const xml = drawingPartXml(snapshot)
  return [...xml.matchAll(new RegExp(`<xdr:${axis}>(\\d+)</xdr:${axis}>`, 'gu'))].map((match) => Number(match[1]))
}

function drawingPartXml(snapshot: WorkbookSnapshot): string {
  const part = snapshot.workbook.metadata?.drawingArtifacts?.parts.find((candidate) => candidate.path === 'xl/drawings/drawing1.xml')
  return part ? new TextDecoder().decode(decodeBase64(part.dataBase64)) : ''
}

function encodedTextPart(text: string): string {
  return encodeBase64(new TextEncoder().encode(text))
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = ''
  for (const byte of bytes) {
    binary += String.fromCharCode(byte)
  }
  return globalThis.btoa(binary)
}

function decodeBase64(dataBase64: string): Uint8Array {
  const binary = globalThis.atob(dataBase64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

const drawingXml =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
  '<xdr:twoCellAnchor editAs="oneCell">' +
  '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/></xdr:nvSpPr></xdr:sp>' +
  '<xdr:clientData/>' +
  '</xdr:twoCellAnchor>' +
  '</xdr:wsDr>'
