import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../index.js'
import type { WorkbookSnapshot } from '@bilig/protocol'

describe('conditional format artifact metadata', () => {
  it('preserves imported advanced conditional format artifacts across headless mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts' })
    await engine.ready()
    engine.importSnapshot(advancedConditionalFormatSnapshot())

    engine.setCellValue('Dashboard', 'D1', 'agent edit')

    const exported = engine.exportSnapshot()
    expect(exported.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml).toContain('type="dataBar"')
    expect(exported.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml).toContain('type="colorScale"')
    expect(exported.sheets[0]?.metadata?.conditionalFormatArtifacts?.xml).toContain('type="iconSet"')

    const restored = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts-restored' })
    await restored.ready()
    restored.importSnapshot(exported)

    expect(restored.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts).toEqual(
      exported.sheets[0]?.metadata?.conditionalFormatArtifacts,
    )
  })

  it('rewrites advanced conditional format artifact ranges across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts-structural' })
    await engine.ready()
    engine.importSnapshot(advancedConditionalFormatSnapshot())

    engine.insertRows('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="A2:A4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="B2:B4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="C2:C4"')

    engine.insertColumns('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="B2:B4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="C2:C4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="D2:D4"')

    engine.deleteRows('Dashboard', 1, 3)
    expect(engine.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts).toBeUndefined()
  })

  it('rewrites namespace-qualified and self-closing conditional format artifacts across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts-qualified-structural' })
    await engine.ready()
    engine.importSnapshot(namespaceQualifiedConditionalFormatSnapshot())

    engine.insertRows('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<x:conditionalFormatting xmlns:x="urn:test" sqref="A2:A4">')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<conditionalFormatting sqref="C2:C4"/>')

    engine.insertColumns('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<x:conditionalFormatting xmlns:x="urn:test" sqref="B2:B4">')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<conditionalFormatting sqref="D2:D4"/>')
  })

  it('rewrites conditional format artifact formulas across structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts-formula-structural' })
    await engine.ready()
    engine.importSnapshot(formulaConditionalFormatArtifactSnapshot())

    engine.insertRows('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="A2:A4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<formula>A2&gt;15</formula>')

    engine.insertColumns('Dashboard', 0, 1)
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="B2:B4"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<formula>B2&gt;15</formula>')
  })

  it('rewrites cross-sheet conditional format formulas across target structural edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'conditional-format-artifacts-cross-sheet-formula-structural' })
    await engine.ready()
    engine.importSnapshot(crossSheetFormulaConditionalFormatArtifactSnapshot())

    engine.insertRows('Inputs', 0, 1)

    expect(engine.exportSnapshot().sheets[0]?.metadata?.conditionalFormats?.[0]?.rule).toEqual({
      kind: 'formula',
      formula: '=Inputs!A2>15',
    })
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('sqref="A1:A3"')
    expect(exportedConditionalFormatArtifactXml(engine)).toContain('<formula>Inputs!A2&gt;15</formula>')
  })
})

function exportedConditionalFormatArtifactXml(engine: SpreadsheetEngine): string {
  const xml = engine.exportSnapshot().sheets[0]?.metadata?.conditionalFormatArtifacts?.xml
  if (!xml) {
    throw new Error('Expected conditional format artifact XML')
  }
  return xml
}

function advancedConditionalFormatSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Advanced conditional format artifacts' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        metadata: {
          conditionalFormatArtifacts: {
            xml: [
              '<conditionalFormatting sqref="A1:A3">',
              '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/>',
              '<color rgb="FF63C384"/></dataBar></cfRule>',
              '</conditionalFormatting>',
              '<conditionalFormatting sqref="B1:B3">',
              '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/>',
              '<cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>',
              '</conditionalFormatting>',
              '<conditionalFormatting sqref="C1:C3">',
              '<cfRule type="iconSet" priority="3"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/>',
              '<cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>',
              '</conditionalFormatting>',
            ].join(''),
          },
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'B1', value: 20 },
          { address: 'C1', value: 30 },
          { address: 'A2', value: 20 },
          { address: 'B2', value: 40 },
          { address: 'C2', value: 60 },
          { address: 'A3', value: 30 },
          { address: 'B3', value: 60 },
          { address: 'C3', value: 90 },
        ],
      },
    ],
  }
}

function namespaceQualifiedConditionalFormatSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Namespace-qualified conditional format artifacts' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        metadata: {
          conditionalFormatArtifacts: {
            xml: [
              '<x:conditionalFormatting xmlns:x="urn:test" sqref="A1:A3">',
              '<x:cfRule type="dataBar" priority="1"><x:dataBar><x:cfvo type="min"/><x:cfvo type="max"/>',
              '<x:color rgb="FF63C384"/></x:dataBar></x:cfRule>',
              '</x:conditionalFormatting>',
              '<conditionalFormatting sqref="C1:C3"/>',
            ].join(''),
          },
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'C1', value: 30 },
          { address: 'A2', value: 20 },
          { address: 'C2', value: 60 },
          { address: 'A3', value: 30 },
          { address: 'C3', value: 90 },
        ],
      },
    ],
  }
}

function formulaConditionalFormatArtifactSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Formula conditional format artifacts' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        metadata: {
          conditionalFormatArtifacts: {
            xml: [
              '<conditionalFormatting sqref="A1:A3">',
              '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;15</formula></cfRule>',
              '</conditionalFormatting>',
            ].join(''),
          },
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
      },
    ],
  }
}

function crossSheetFormulaConditionalFormatArtifactSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Cross-sheet formula conditional format artifacts' },
    sheets: [
      {
        id: 1,
        name: 'Dashboard',
        order: 0,
        metadata: {
          conditionalFormats: [
            {
              id: 'cross-sheet-formula-highlight',
              range: { sheetName: 'Dashboard', startAddress: 'A1', endAddress: 'A3' },
              rule: { kind: 'formula', formula: '=Inputs!A1>15' },
              style: { fill: { backgroundColor: '#ffeb84' } },
              priority: 1,
            },
          ],
          conditionalFormatArtifacts: {
            xml: [
              '<conditionalFormatting sqref="A1:A3">',
              '<cfRule type="expression" dxfId="0" priority="1"><formula>Inputs!A1&gt;15</formula></cfRule>',
              '</conditionalFormatting>',
            ].join(''),
          },
        },
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
      },
      {
        id: 2,
        name: 'Inputs',
        order: 1,
        cells: [
          { address: 'A1', value: 10 },
          { address: 'A2', value: 20 },
          { address: 'A3', value: 30 },
        ],
      },
    ],
  }
}
