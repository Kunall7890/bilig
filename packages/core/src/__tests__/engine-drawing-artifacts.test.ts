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
          dataBase64: 'AQID',
          byteLength: 3,
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
              dataBase64: 'AQID',
              byteLength: 3,
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
