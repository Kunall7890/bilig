import { Buffer } from 'node:buffer'

import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

const controlPropertiesXml = '<formControlPr objectType="Button" lockText="1"/>'
const vmlDrawingXml = [
  '<xml xmlns:x="urn:schemas-microsoft-com:office:excel">',
  '<x:ClientData ObjectType="Button"><x:Anchor>6, 11, 3, 5, 11, 7, 4, 12</x:Anchor></x:ClientData>',
  '</xml>',
].join('')

const workbookControlArtifacts = {
  parts: [
    encodePackagePart('xl/ctrlProps/ctrlProp1.xml', controlPropertiesXml),
    encodePackagePart('xl/drawings/vmlDrawing1.vml', vmlDrawingXml),
  ],
  contentTypeDefaults: [{ extension: 'vml', contentType: 'application/vnd.openxmlformats-officedocument.vmlDrawing' }],
  contentTypeOverrides: [{ partName: '/xl/ctrlProps/ctrlProp1.xml', contentType: 'application/vnd.ms-excel.controlproperties+xml' }],
}

const sheetControlArtifacts = {
  worksheetRootOpenTag:
    '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14">',
  controlsXml: [
    '<controls><control shapeId="1025" r:id="rId4" name="Button 1">',
    '<controlPr defaultSize="0" macro="[0]!WriteHelloWorld">',
    '<anchor moveWithCells="1">',
    '<from><xdr:col>6</xdr:col><xdr:colOff>142875</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>66675</xdr:rowOff></from>',
    '<to><xdr:col>11</xdr:col><xdr:colOff>85725</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>152400</xdr:rowOff></to>',
    '</anchor>',
    '</controlPr>',
    '</control></controls>',
  ].join(''),
  relationships: [
    {
      id: 'rId3',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
      target: '../drawings/vmlDrawing1.vml',
    },
    {
      id: 'rId4',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp',
      target: '../ctrlProps/ctrlProp1.xml',
    },
  ],
}

describe('engine worksheet control artifacts', () => {
  it('roundtrips workbook and worksheet control artifacts through runtime snapshots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'control-artifacts-runtime' })
    await engine.ready()

    engine.importSnapshot(controlArtifactsSnapshot())
    engine.setCellValue('Model', 'B1', 'headless edit')

    expect(engine.exportSnapshot().workbook.metadata?.controlArtifacts).toEqual(workbookControlArtifacts)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.controlArtifacts).toEqual(sheetControlArtifacts)

    const restored = new SpreadsheetEngine({ workbookName: 'control-artifacts-restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.exportSnapshot().workbook.metadata?.controlArtifacts).toEqual(workbookControlArtifacts)
    expect(restored.exportSnapshot().sheets[0]?.metadata?.controlArtifacts).toEqual(sheetControlArtifacts)
  })

  it('rewrites control DrawingML and VML anchors across structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'control-artifacts-structure' })
    await engine.ready()

    engine.importSnapshot(controlArtifactsSnapshot())
    engine.insertRows('Model', 0, 1)

    const exported = engine.exportSnapshot()
    const controlsXml = exported.sheets[0]?.metadata?.controlArtifacts?.controlsXml ?? ''
    const vmlXml = decodePackagePart(
      exported.workbook.metadata?.controlArtifacts?.parts.find((part) => part.path === 'xl/drawings/vmlDrawing1.vml')?.dataBase64 ?? '',
    )

    expect(controlsXml).toContain('<xdr:row>4</xdr:row>')
    expect(controlsXml).toContain('<xdr:row>5</xdr:row>')
    expect(vmlXml).toContain('<x:Anchor>6, 11, 4, 5, 11, 7, 5, 12</x:Anchor>')
  })
})

function controlArtifactsSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'Control artifacts runtime',
      metadata: { controlArtifacts: workbookControlArtifacts },
    },
    sheets: [
      {
        id: 1,
        name: 'Model',
        order: 0,
        metadata: { controlArtifacts: sheetControlArtifacts },
        cells: [{ address: 'A1', value: 'Control fixture' }],
      },
    ],
  }
}

function encodePackagePart(
  path: string,
  xml: string,
): {
  readonly path: string
  readonly storage: 'base64'
  readonly dataBase64: string
  readonly byteLength: number
} {
  const bytes = Buffer.from(xml, 'utf8')
  return {
    path,
    storage: 'base64',
    dataBase64: bytes.toString('base64'),
    byteLength: bytes.byteLength,
  }
}

function decodePackagePart(dataBase64: string): string {
  return Buffer.from(dataBase64, 'base64').toString('utf8')
}
