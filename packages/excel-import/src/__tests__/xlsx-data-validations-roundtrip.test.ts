import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { SpreadsheetEngine } from '@bilig/core'

import { exportXlsx, importXlsx } from '../index.js'

interface DataValidationSummary {
  readonly type: string | null
  readonly sqref: string | null
  readonly promptTitle: string | null
  readonly prompt: string | null
  readonly formula1: string
  readonly formula2: string
}

describe('data validation roundtrip', () => {
  it('preserves same-sheet list ranges and prompt-only validations across XLSX round trips', () => {
    const source = buildDataValidationWorkbookBytes()
    const imported = importXlsx(source, 'data-validations.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual([
      {
        range: { sheetName: 'Model', startAddress: 'B2', endAddress: 'B2' },
        rule: { kind: 'list', source: { kind: 'range-ref', sheetName: 'Model', startAddress: 'D2', endAddress: 'D4' } },
        allowBlank: true,
      },
      {
        range: { sheetName: 'Model', startAddress: 'B3', endAddress: 'B3' },
        rule: { kind: 'any' },
        allowBlank: true,
        promptTitle: 'Use model choices',
        promptMessage: 'Pick a case from the list.',
      },
      {
        range: { sheetName: 'Model', startAddress: 'B4', endAddress: 'B4' },
        rule: { kind: 'decimal', operator: 'between', values: [0, 1] },
        allowBlank: true,
        promptTitle: 'Debt ratio',
        promptMessage: 'Enter a ratio from 0 to 1.',
      },
    ])

    expect(readDataValidations(exportXlsx(imported.snapshot))).toEqual([
      {
        type: 'list',
        sqref: 'B2',
        promptTitle: null,
        prompt: null,
        formula1: '$D$2:$D$4',
        formula2: '',
      },
      {
        type: null,
        sqref: 'B3',
        promptTitle: 'Use model choices',
        prompt: 'Pick a case from the list.',
        formula1: '',
        formula2: '',
      },
      {
        type: 'decimal',
        sqref: 'B4',
        promptTitle: 'Debt ratio',
        prompt: 'Enter a ratio from 0 to 1.',
        formula1: '0',
        formula2: '1',
      },
    ])
  })

  it('escapes structured-reference list validation sources with special table headers', () => {
    const snapshot = {
      version: 1,
      workbook: {
        name: 'Structured Validation',
        metadata: {
          tables: [
            {
              name: 'Sales',
              sheetName: 'Model',
              startAddress: 'D1',
              endAddress: 'D3',
              columnNames: ['# Units'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Model',
          order: 0,
          cells: [
            { address: 'A1', value: 'Choice' },
            { address: 'D1', value: '# Units' },
            { address: 'D2', value: 'Base' },
            { address: 'D3', value: 'Upside' },
          ],
          metadata: {
            validations: [
              {
                range: { sheetName: 'Model', startAddress: 'A2', endAddress: 'A2' },
                rule: { kind: 'list', source: { kind: 'structured-ref', tableName: 'Sales', columnName: '# Units' } },
              },
            ],
          },
        },
      ],
    } as const

    const exported = exportXlsx(snapshot)
    expect(readDataValidations(exported)).toEqual([
      {
        type: 'list',
        sqref: 'A2',
        promptTitle: null,
        prompt: null,
        formula1: 'Sales[&apos;# Units]',
        formula2: '',
      },
    ])

    const imported = importXlsx(exported, 'structured-validation.xlsx')
    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual(snapshot.sheets[0]?.metadata?.validations)
  })

  it('preserves formula list validation sources across XLSX import and export', () => {
    const snapshot = {
      version: 1,
      workbook: { name: 'Formula Validation' },
      sheets: [
        {
          id: 1,
          name: 'Model',
          order: 0,
          cells: [
            { address: 'A1', value: 'Choice' },
            { address: 'D2', value: 'Base' },
            { address: 'D3', value: 'Downside' },
            { address: 'D4', value: 'Upside' },
          ],
          metadata: {
            validations: [
              {
                range: { sheetName: 'Model', startAddress: 'A2', endAddress: 'A2' },
                rule: { kind: 'list', source: { kind: 'formula', formula: '=OFFSET($D$2,0,0,3,1)' } },
              },
            ],
          },
        },
      ],
    } as const

    const exported = exportXlsx(snapshot)
    expect(readDataValidations(exported)).toEqual([
      {
        type: 'list',
        sqref: 'A2',
        promptTitle: null,
        prompt: null,
        formula1: 'OFFSET($D$2,0,0,3,1)',
        formula2: '',
      },
    ])

    const imported = importXlsx(exported, 'formula-validation.xlsx')
    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual(snapshot.sheets[0]?.metadata?.validations)
  })

  it('preserves broken list validation references during XLSX import', async () => {
    const source = buildBrokenListValidationWorkbookBytes()
    const imported = importXlsx(source, 'broken-list-validation.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual([
      {
        range: { sheetName: 'Model', startAddress: 'A1', endAddress: 'A1' },
        rule: { kind: 'list', source: { kind: 'formula', formula: '=#REF!' } },
        allowBlank: true,
      },
    ])
    expect(readDataValidations(exportXlsx(imported.snapshot))).toEqual([
      {
        type: 'list',
        sqref: 'A1',
        promptTitle: null,
        prompt: null,
        formula1: '#REF!',
        formula2: '',
      },
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'broken-list-validation-import' })
    await engine.ready()

    expect(() => engine.importSnapshot(imported.snapshot)).not.toThrow()
  })

  it('keeps engine table-sort validations anchored across XLSX export and reimport', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'table-sort-validation-roundtrip' })
    await engine.ready()
    engine.importSnapshot({
      version: 1,
      workbook: {
        name: 'Table sort validation roundtrip',
        metadata: {
          tables: [
            {
              name: 'Sales',
              sheetName: 'Ledger',
              startAddress: 'A1',
              endAddress: 'D6',
              columnNames: ['Region', 'Amount', 'Invoice', 'Double'],
              columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Invoice' }, { name: 'Double' }],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Ledger',
          order: 0,
          metadata: {
            validations: [
              {
                range: { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' },
                rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
              },
            ],
          },
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Amount' },
            { address: 'C1', value: 'Invoice' },
            { address: 'D1', value: 'Double' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 10 },
            { address: 'C2', value: 'invoice-001' },
            { address: 'D2', formula: 'B2*2', value: 20 },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 40 },
            { address: 'C3', value: 'invoice-002' },
            { address: 'D3', formula: 'B3*2', value: 80 },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 30 },
            { address: 'C4', value: 'invoice-003' },
            { address: 'D4', formula: 'B4*2', value: 60 },
            { address: 'A5', value: 'West' },
            { address: 'B5', value: 20 },
            { address: 'C5', value: 'invoice-004' },
            { address: 'D5', formula: 'B5*2', value: 40 },
            { address: 'A6', value: 'East' },
            { address: 'B6', value: 50 },
            { address: 'C6', value: 'invoice-005' },
            { address: 'D6', formula: 'B6*2', value: 100 },
          ],
        },
      ],
    })

    expect(engine.sortTable('Ledger', 'Sales', [{ keyAddress: 'B1', direction: 'desc' }])).toBe(true)
    expect(engine.getDataValidation('Ledger', { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' })).toMatchObject({
      rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
    })

    const exported = exportXlsx(engine.exportSnapshot())
    expect(readDataValidations(exported)).toEqual([
      {
        type: 'whole',
        sqref: 'B6',
        promptTitle: null,
        prompt: null,
        formula1: '0',
        formula2: '',
      },
    ])

    const imported = importXlsx(exported, 'table-sort-validation-roundtrip.xlsx')
    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.tables?.[0]?.sortState).toBe(
      '<sortState ref="A2:D6"><sortCondition descending="1" ref="B2:B6"/></sortState>',
    )
    expect(imported.snapshot.sheets[0]?.metadata?.validations).toEqual([
      {
        range: { sheetName: 'Ledger', startAddress: 'B6', endAddress: 'B6' },
        rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
      },
    ])
  })
})

function buildDataValidationWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Input', 'Value', '', 'Choices'],
      ['Scenario', '', '', 'Base'],
      ['Guidance', '', '', 'Downside'],
      ['Debt ratio', '', '', 'Upside'],
    ]),
    'Model',
  )
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(
    sheetXml.replace(
      '</worksheet>',
      '<dataValidations count="3">' +
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2"><formula1>$D$2:$D$4</formula1></dataValidation>' +
        '<dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Use model choices" prompt="Pick a case from the list." sqref="B3"/>' +
        '<dataValidation type="decimal" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Debt ratio" prompt="Enter a ratio from 0 to 1." sqref="B4"><formula1>0</formula1><formula2>1</formula2></dataValidation>' +
        '</dataValidations></worksheet>',
    ),
  )
  return zipSync(zip)
}

function buildBrokenListValidationWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Choice']]), 'Model')
  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  zip[sheetPath] = strToU8(
    sheetXml.replace(
      '</worksheet>',
      '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A1"><formula1>#REF!</formula1></dataValidation></dataValidations></worksheet>',
    ),
  )
  return zipSync(zip)
}

function readDataValidations(bytes: Uint8Array): DataValidationSummary[] {
  const sheetXml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  return [...sheetXml.matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/gu)].map((match) => {
    const attributes = match[1] ?? ''
    const body = match[2] ?? ''
    return {
      type: readAttribute(attributes, 'type'),
      sqref: readAttribute(attributes, 'sqref'),
      promptTitle: readAttribute(attributes, 'promptTitle'),
      prompt: readAttribute(attributes, 'prompt'),
      formula1: readFormula(body, 'formula1'),
      formula2: readFormula(body, 'formula2'),
    }
  })
}

function readFormula(body: string, elementName: 'formula1' | 'formula2'): string {
  return new RegExp(`<${elementName}>([\\s\\S]*?)</${elementName}>`, 'u').exec(body)?.[1] ?? ''
}

function readAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}
