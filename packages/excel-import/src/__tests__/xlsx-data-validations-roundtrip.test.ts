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

  it('ignores broken list validation references during XLSX import', async () => {
    const source = buildBrokenListValidationWorkbookBytes()
    const imported = importXlsx(source, 'broken-list-validation.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.validations ?? []).toEqual([])

    const engine = new SpreadsheetEngine({ workbookName: 'broken-list-validation-import' })
    await engine.ready()

    expect(() => engine.importSnapshot(imported.snapshot)).not.toThrow()
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
