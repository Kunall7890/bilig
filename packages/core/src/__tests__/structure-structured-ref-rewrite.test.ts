import { describe, expect, it } from 'vitest'

import {
  rewriteFormulaSourceForDeletedStructuredReferences,
  rewriteFormulaSourceForRenamedStructuredReference,
} from '../engine/services/structure-structured-ref-rewrite.js'

describe('structured reference formula rewrites', () => {
  it('returns undefined when structured references are absent or unaffected', () => {
    expect(rewriteFormulaSourceForDeletedStructuredReferences('A1+1', [])).toBeUndefined()
    expect(
      rewriteFormulaSourceForDeletedStructuredReferences('SUM(Sales[Amount])', [
        {
          tableName: 'Other',
          columnName: 'Amount',
        },
      ]),
    ).toBeUndefined()
    expect(
      rewriteFormulaSourceForRenamedStructuredReference('SUM(Sales[Amount])', {
        tableName: 'Sales',
        oldColumnName: 'Margin',
        newColumnName: 'Revenue',
      }),
    ).toBeUndefined()
  })

  it('rewrites deleted structured references through nested expressions', () => {
    const deleted = [{ tableName: ' sales ', columnName: ' amount ' }]

    expect(rewriteFormulaSourceForDeletedStructuredReferences('SUM(Sales[Amount])', deleted)).toBe('SUM(#REF!)')
    expect(rewriteFormulaSourceForDeletedStructuredReferences('-Sales[Amount]', deleted)).toBe('-#REF!')
    expect(rewriteFormulaSourceForDeletedStructuredReferences('Sales[Amount]+Sales[Tax]', deleted)).toBe('#REF!+Sales[Tax]')
    expect(rewriteFormulaSourceForDeletedStructuredReferences('{Sales[Amount],1;2,Sales[Tax]}', deleted)).toBe('{#REF!,1;2,Sales[Tax]}')
  })

  it('rewrites renamed structured references through nested expressions', () => {
    const renamed = {
      tableName: 'sales',
      oldColumnName: 'amount',
      newColumnName: 'Revenue',
    }

    expect(rewriteFormulaSourceForRenamedStructuredReference('SUM(Sales[Amount])', renamed)).toBe('SUM(Sales[Revenue])')
    expect(rewriteFormulaSourceForRenamedStructuredReference('-Sales[Amount]', renamed)).toBe('-Sales[Revenue]')
    expect(rewriteFormulaSourceForRenamedStructuredReference('Sales[Amount]+Sales[Tax]', renamed)).toBe('Sales[Revenue]+Sales[Tax]')
    expect(rewriteFormulaSourceForRenamedStructuredReference('{Sales[Amount],1;2,Sales[Tax]}', renamed)).toBe(
      '{Sales[Revenue],1;2,Sales[Tax]}',
    )
  })
})
