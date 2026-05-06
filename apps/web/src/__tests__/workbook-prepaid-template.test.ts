import { describe, expect, it } from 'vitest'
import { compileFormula } from '@bilig/formula'
import { buildPrepaidAmortizationTemplateMutations } from '../workbook-prepaid-template.js'

describe('prepaid amortization workbook template', () => {
  it('builds a richer daily-prorated template than the reference sheet', () => {
    const mutations = buildPrepaidAmortizationTemplateMutations('Prepaid Schedule')
    const renderCommit = mutations.find((mutation) => mutation.method === 'renderCommit')

    expect(renderCommit).toBeDefined()
    expect(renderCommit?.args[0]).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          addr: 'A1',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
          value: 'Prepaid Amortization Schedule',
        }),
        expect.objectContaining({
          addr: 'A5',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
          value: 'Vendor',
        }),
        expect.objectContaining({
          addr: 'A6',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
          value: 'ABC Insurance',
        }),
        expect.objectContaining({
          addr: 'A10',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
          value: 'Cybersecurity Policy',
        }),
        expect.objectContaining({
          addr: 'H6',
          formula: 'ROUND(IFERROR($E6*MAX(0,MIN($D6,EOMONTH(DATE(2024,1,1),0))-MAX($C6,DATE(2024,1,1))+1)/($D6-$C6+1),0),2)',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
        }),
        expect.objectContaining({
          addr: 'V6',
          formula:
            'IF(COUNTA(A6:E6)=0,"",IF(OR(C6="",D6=""),"Missing dates",IF(E6<=0,"Missing amount",IF(D6<C6,"Check dates",IF(U6<0,"Over-amortized",IF(U6<=0,"Complete",IF(T6=0,"Not started","In progress")))))))',
          kind: 'upsertCell',
          sheetName: 'Prepaid Schedule',
        }),
      ]),
    )
    const renderCommitOps = renderCommit?.args[0]
    expect(Array.isArray(renderCommitOps)).toBe(true)
    if (!Array.isArray(renderCommitOps)) {
      throw new Error('Expected renderCommit ops')
    }
    expect(renderCommitOps.filter((op) => JSON.stringify(op).includes('formula')).length).toBeGreaterThan(70)
    expect(mutations).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ method: 'mergeCells', args: [{ sheetName: 'Prepaid Schedule', startAddress: 'A1', endAddress: 'W1' }] }),
        expect.objectContaining({ method: 'mergeCells', args: [{ sheetName: 'Prepaid Schedule', startAddress: 'A4', endAddress: 'W4' }] }),
        expect.objectContaining({ method: 'setFreezePane', args: ['Prepaid Schedule', 5, 0] }),
        expect.objectContaining({ method: 'updateColumnMetadata', args: ['Prepaid Schedule', 0, 1, 184, null] }),
        expect.objectContaining({
          method: 'setRangeNumberFormat',
          args: [
            { sheetName: 'Prepaid Schedule', startAddress: 'C6', endAddress: 'D10' },
            { kind: 'date', dateStyle: 'iso' },
          ],
        }),
        expect.objectContaining({
          method: 'setRangeStyle',
          args: [
            { sheetName: 'Prepaid Schedule', startAddress: 'A5', endAddress: 'W5' },
            expect.objectContaining({
              fill: { backgroundColor: '#21563A' },
              font: expect.objectContaining({ bold: true, color: '#FFFFFF' }),
            }),
          ],
        }),
      ]),
    )
  })

  it('builds the template from custom prepaid parameters', () => {
    const mutations = buildPrepaidAmortizationTemplateMutations('Custom Prepaids', {
      year: 2026,
      dataRowCount: 6,
      items: [
        {
          vendor: 'TenantWorks',
          description: 'Facilities platform',
          start: [2026, 2, 1],
          end: [2026, 7, 31],
          amount: 6_600,
          notes: 'Custom schedule item',
        },
      ],
    })
    const renderCommit = mutations.find((mutation) => mutation.method === 'renderCommit')
    const renderCommitOps = renderCommit?.args[0]
    expect(Array.isArray(renderCommitOps)).toBe(true)
    if (!Array.isArray(renderCommitOps)) {
      throw new Error('Expected renderCommit ops')
    }

    expect(renderCommitOps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          addr: 'H5',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
          value: 'Jan 2026',
        }),
        expect.objectContaining({
          addr: 'T5',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
          value: '2026 Amortized',
        }),
        expect.objectContaining({
          addr: 'A6',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
          value: 'TenantWorks',
        }),
        expect.objectContaining({
          addr: 'H6',
          formula: 'ROUND(IFERROR($E6*MAX(0,MIN($D6,EOMONTH(DATE(2026,1,1),0))-MAX($C6,DATE(2026,1,1))+1)/($D6-$C6+1),0),2)',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
        }),
        expect.objectContaining({
          addr: 'H11',
          formula: 'ROUND(IFERROR($E11*MAX(0,MIN($D11,EOMONTH(DATE(2026,1,1),0))-MAX($C11,DATE(2026,1,1))+1)/($D11-$C11+1),0),2)',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
        }),
        expect.objectContaining({
          addr: 'V11',
          formula:
            'IF(COUNTA(A11:E11)=0,"",IF(OR(C11="",D11=""),"Missing dates",IF(E11<=0,"Missing amount",IF(D11<C11,"Check dates",IF(U11<0,"Over-amortized",IF(U11<=0,"Complete",IF(T11=0,"Not started","In progress")))))))',
          kind: 'upsertCell',
          sheetName: 'Custom Prepaids',
        }),
      ]),
    )
    expect(renderCommitOps).not.toEqual(expect.arrayContaining([expect.objectContaining({ value: 'ABC Insurance' })]))
    const blankRowStatus = renderCommitOps.find((op) => 'addr' in op && op.addr === 'V11')
    expect(blankRowStatus && 'formula' in blankRowStatus ? compileFormula(blankRowStatus.formula).astMatchesSource : false).toBe(true)
    expect(blankRowStatus && 'formula' in blankRowStatus ? compileFormula(blankRowStatus.formula).deps : []).toEqual([
      'A11:E11',
      'C11',
      'D11',
      'E11',
      'U11',
      'T11',
    ])
    expect(mutations).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ method: 'clearRange', args: [{ sheetName: 'Custom Prepaids', startAddress: 'A1', endAddress: 'W40' }] }),
        expect.objectContaining({
          method: 'setRangeNumberFormat',
          args: [
            { sheetName: 'Custom Prepaids', startAddress: 'C6', endAddress: 'D11' },
            { kind: 'date', dateStyle: 'iso' },
          ],
        }),
        expect.objectContaining({
          method: 'setRangeStyle',
          args: [
            { sheetName: 'Custom Prepaids', startAddress: 'A6', endAddress: 'W11' },
            expect.objectContaining({
              font: expect.objectContaining({ color: '#1F2933' }),
            }),
          ],
        }),
      ]),
    )
  })
})
