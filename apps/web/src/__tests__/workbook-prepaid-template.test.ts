import { describe, expect, it } from 'vitest'
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
          formula: 'IF(U6<=0,"Complete",IF(T6=0,"Not started","In progress"))',
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
})
