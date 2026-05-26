import Ajv2020 from 'ajv/dist/2020.js'
import { describe, expect, it } from 'vitest'
import { checkWorkbookRunResultDescription, workbookJsonSchemas } from '../index.js'

function validateRunResultSchema(payload: unknown): boolean {
  const ajv = new Ajv2020({ allErrors: true, strict: false })
  return Boolean(ajv.compile(workbookJsonSchemas.runResult)(payload))
}

function runDescriptionWithNoopEffect(
  commandKind: string,
  effect: Record<string, unknown>,
  options: { readonly proofCommandKind?: string } = {},
): Record<string, unknown> {
  const proofCommandKind = options.proofCommandKind ?? commandKind
  return {
    status: 'done',
    apply: {
      matched: true,
      commandReceipts: [
        {
          commandIndex: 0,
          commandKind,
          commandDigest: 'bilig-command-v1:schema',
          previewOps: [],
          appliedOps: [],
          noop: {
            reason: 'already_satisfied',
            proof: {
              source: 'schema-test',
              evidence: 'adapter_zero_ops',
              commandKind: proofCommandKind,
              commandDigest: 'bilig-command-v1:schema',
              opCount: 0,
              effect,
            },
          },
        },
      ],
    },
    changed: [],
    checks: [],
  }
}

describe('@bilig/workbook run-result schema api', () => {
  it('keeps no-op effect proof schemas aligned with the description checker', () => {
    const invalidDescriptions = [
      runDescriptionWithNoopEffect('writeValue', { kind: 'writeValue' }),
      runDescriptionWithNoopEffect('writeFormula', { kind: 'writeFormula' }),
      runDescriptionWithNoopEffect('clear', { kind: 'clear' }),
      runDescriptionWithNoopEffect('format', { kind: 'format' }),
      runDescriptionWithNoopEffect('writeValue', { kind: 'writeFormula', formula: 'SUM(A1:A3)' }),
      runDescriptionWithNoopEffect('writeValue', { kind: 'writeValue', value: 12 }, { proofCommandKind: 'writeFormula' }),
      runDescriptionWithNoopEffect('op', { kind: 'op', opKind: 'setCellValue' }),
      runDescriptionWithNoopEffect('op', {
        kind: 'op',
        opKind: 'setCellFormula',
        op: { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 12 },
      }),
    ] as const

    for (const description of invalidDescriptions) {
      expect(validateRunResultSchema(description), JSON.stringify(description)).toBe(false)
      expect(checkWorkbookRunResultDescription(description).status, JSON.stringify(description)).toBe('invalid')
    }

    const validDescriptions = [
      runDescriptionWithNoopEffect('writeValue', { kind: 'writeValue', value: 12 }),
      runDescriptionWithNoopEffect('writeFormula', { kind: 'writeFormula', formula: 'SUM(A1:A3)' }),
      runDescriptionWithNoopEffect('clear', { kind: 'clear', cleared: true }),
      runDescriptionWithNoopEffect('format', { kind: 'format', numberFormat: '0.00' }),
      runDescriptionWithNoopEffect('op', {
        kind: 'op',
        opKind: 'setCellValue',
        op: { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 12 },
      }),
    ] as const

    for (const description of validDescriptions) {
      expect(validateRunResultSchema(description), JSON.stringify(description)).toBe(true)
      expect(checkWorkbookRunResultDescription(description).status, JSON.stringify(description)).toBe('valid')
    }
  })
})
