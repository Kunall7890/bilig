import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { checkWorkbookReadbackProof, isWorkbookReadbackProof } from '../index.js'

const fixtureRoot = new URL('../../fixtures/', import.meta.url)

function readFixture(name: string): unknown {
  return JSON.parse(readFileSync(new URL(name, fixtureRoot), 'utf8'))
}

function requiredRecord(value: unknown, path: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${path} must be an object`)
  }
  return value
}

function requiredArray(value: unknown, path: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`${path} must be an array`)
  }
  return value
}

describe('@bilig/workbook readback proof api', () => {
  it('validates transported readback proof into frozen verified checks', () => {
    const proof = readFixture('readback-proof.json')
    const result = checkWorkbookReadbackProof(proof)

    expect(result.status).toBe('valid')
    if (result.status !== 'valid') {
      throw new Error('readback-proof fixture failed validation')
    }
    expect(isWorkbookReadbackProof(proof)).toBe(true)
    expect(Object.isFrozen(result)).toBe(true)
    expect(Object.isFrozen(result.proof)).toBe(true)
    expect(Object.isFrozen(result.proof.checks)).toBe(true)
    expect(Object.isFrozen(result.proof.readbacks)).toBe(true)
    expect(result.verification.status).toBe('passed')
    expect(result.proof.checks).toEqual([
      expect.objectContaining({
        status: 'passed',
        kind: 'formulaEquals',
        proof: expect.objectContaining({
          source: 'readback',
          formula: 'input*factor',
          expectedFormula: 'input*factor',
          materializedFormula: 'input*factor',
        }),
      }),
    ])
  })

  it('accepts transported nested refs without live helper methods', () => {
    const table = {
      kind: 'table',
      id: 'table_generic',
      label: 'generic table',
      headers: ['key', 'result'],
    }
    const resultColumn = {
      kind: 'column',
      id: 'column_result',
      label: 'result',
      table,
      name: 'result',
    }

    const result = checkWorkbookReadbackProof({
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: resultColumn,
          message: 'result equals 7',
          expectation: {
            kind: 'valueEquals',
            value: 7,
          },
        },
      ],
      readbacks: [
        {
          target: resultColumn,
          value: 7,
        },
      ],
    })

    expect(result.status).toBe('valid')
    if (result.status !== 'valid') {
      throw new Error('transported nested readback proof failed validation')
    }
    expect(result.proof.checks[0]?.target).toEqual(expect.objectContaining({ kind: 'column', name: 'result' }))
    expect(result.proof.checks[0]?.proof).toEqual({
      source: 'readback',
      value: 7,
    })
  })

  it('returns stable issues when readbacks do not prove checks', () => {
    const proof = requiredRecord(readFixture('readback-proof.json'), 'proof')
    const readbacks = requiredArray(proof['readbacks'], 'proof.readbacks')
    requiredRecord(readbacks[0], 'proof.readbacks[0]')['formula'] = 'input+factor'

    const result = checkWorkbookReadbackProof(proof)

    expect(result.status).toBe('invalid')
    if (result.status !== 'invalid') {
      throw new Error('mismatched readback proof unexpectedly passed')
    }
    expect(isWorkbookReadbackProof(proof)).toBe(false)
    expect(result.issues).toEqual([
      expect.objectContaining({
        code: 'formula_mismatch',
        message: 'result expected formula input*factor but read input+factor',
      }),
    ])
    expect(result.verification?.checks).toEqual([
      expect.objectContaining({
        status: 'failed',
        kind: 'formulaEquals',
      }),
    ])
  })

  it('rejects accessor-backed proof roots without invoking getters', () => {
    let invoked = false
    const proof = {
      readbacks: [],
    }
    Object.defineProperty(proof, 'checks', {
      enumerable: true,
      get() {
        invoked = true
        return []
      },
    })

    const result = checkWorkbookReadbackProof(proof)

    expect(invoked).toBe(false)
    expect(result.status).toBe('invalid')
    if (result.status !== 'invalid') {
      throw new Error('accessor-backed readback proof unexpectedly passed')
    }
    expect(result.issues).toEqual([
      expect.objectContaining({
        code: 'readback_invalid',
        message: 'Workbook readback proof at proof.checks must be a data property',
      }),
    ])
  })
})
