import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  checkPlanData,
  checkWorkbookCommandBundle,
  checkWorkbookCommandResultForBundle,
  checkWorkbookReadbackProof,
  checkWorkbookRunResultDescription,
  workbookJsonSchemaBundleHash,
  workbookJsonSchemaHash,
  workbookJsonSchemaHashes,
  workbookJsonSchemaNames,
  workbookJsonSchemas,
  workbookJsonSchemaVersion,
  workbookPlanId,
  workbookRunErrorCodes,
} from '../index.js'

const fixtureRoot = new URL('../../fixtures/', import.meta.url)

function readFixture(name: string): unknown {
  return JSON.parse(readFileSync(new URL(name, fixtureRoot), 'utf8'))
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function objectEntry(value: unknown, key: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`expected object before reading ${key}`)
  }
  const entry = value[key]
  if (!isRecord(entry)) {
    throw new Error(`expected ${key} to be an object`)
  }
  return entry
}

describe('@bilig/workbook schema api', () => {
  it('exports frozen JSON schema artifacts with deterministic hashes', () => {
    expect(workbookJsonSchemaVersion).toBe('bilig-workbook-json-schema-v1')
    expect(workbookJsonSchemaNames).toEqual(['refData', 'planData', 'commandBundle', 'commandResult', 'runResult', 'readbackProof'])
    expect(Object.isFrozen(workbookJsonSchemaNames)).toBe(true)
    expect(Object.isFrozen(workbookJsonSchemas)).toBe(true)
    expect(Object.isFrozen(workbookJsonSchemaHashes)).toBe(true)

    for (const name of workbookJsonSchemaNames) {
      const schema = workbookJsonSchemas[name]
      expect(schema['$schema']).toBe('https://json-schema.org/draft/2020-12/schema')
      expect(schema['$id']).toBe(`https://bilig.proompteng.ai/schemas/workbook/${workbookJsonSchemaVersion}/${name}.schema.json`)
      expect(workbookJsonSchemaHashes[name]).toBe(workbookJsonSchemaHash(schema))
      expect(workbookJsonSchemaHashes[name]).toMatch(/^bilig-schema-v1:[0-9a-f]{32}$/u)
      expect(JSON.parse(JSON.stringify(schema))).toEqual(schema)
    }

    expect(workbookJsonSchemaBundleHash).toBe(
      workbookJsonSchemaHash({
        version: workbookJsonSchemaVersion,
        schemas: workbookJsonSchemas,
      }),
    )
  })

  it('describes run-result proof fields without opaque apply/undo placeholders', () => {
    const defs = objectEntry(workbookJsonSchemas.runResult, '$defs')
    const apply = objectEntry(defs, 'apply')
    const applyProperties = objectEntry(apply, 'properties')
    const commandReceipt = objectEntry(defs, 'applyCommandReceipt')
    const commandReceiptProperties = objectEntry(commandReceipt, 'properties')
    const errorProperties = objectEntry(objectEntry(defs, 'error'), 'properties')

    expect(apply).toMatchObject({
      type: 'object',
      required: ['matched'],
      additionalProperties: false,
    })
    expect(applyProperties['commandReceipts']).toEqual({
      type: 'array',
      items: { $ref: '#/$defs/applyCommandReceipt' },
    })
    expect(applyProperties['baseRevision']).toEqual({ type: 'integer', minimum: 0 })
    expect(applyProperties['revision']).toEqual({ type: 'integer', minimum: 0 })

    expect(commandReceipt).toMatchObject({
      type: 'object',
      required: ['commandIndex', 'commandKind', 'commandDigest', 'previewOps', 'appliedOps'],
      additionalProperties: false,
    })
    expect(commandReceiptProperties['commandIndex']).toEqual({ type: 'integer', minimum: 0 })
    expect(commandReceiptProperties['formulaLabels']).toEqual({
      type: 'array',
      items: {
        type: 'object',
        required: ['name', 'source'],
        additionalProperties: false,
        properties: {
          name: { type: 'string', minLength: 1 },
          source: { type: 'string', minLength: 1 },
        },
      },
    })

    expect(objectEntry(defs, 'undo')).toEqual({
      type: 'object',
      required: ['id'],
      additionalProperties: false,
      properties: {
        id: { type: 'string', minLength: 1 },
        ops: { type: 'array', items: { $ref: '#/$defs/engineOp' } },
      },
    })
    expect(objectEntry(defs, 'unverified')).toEqual({
      type: 'object',
      required: ['kind', 'message'],
      additionalProperties: false,
      properties: {
        kind: { enum: ['apply', 'plan'] },
        message: { type: 'string', minLength: 1 },
      },
    })
    expect(errorProperties['code']).toEqual({ enum: workbookRunErrorCodes })
  })

  it('keeps checked-in contract fixtures aligned with public validators', () => {
    const validPlan = readFixture('valid-plan.json')
    const validPlanCheck = checkPlanData(validPlan)
    expect(validPlanCheck.status).toBe('valid')
    if (validPlanCheck.status !== 'valid') {
      throw new Error('valid-plan fixture failed validation')
    }
    expect(workbookPlanId(validPlanCheck.plan)).toMatch(/^bilig-plan-v1:[0-9a-f]{32}$/u)

    const invalidPlanCheck = checkPlanData(readFixture('invalid-plan.json'))
    expect(invalidPlanCheck.status).toBe('invalid')

    const bundleCheck = checkWorkbookCommandBundle(readFixture('command-bundle.json'))
    expect(bundleCheck.status).toBe('valid')
    if (bundleCheck.status !== 'valid') {
      throw new Error('command-bundle fixture failed validation')
    }

    const commandResultCheck = checkWorkbookCommandResultForBundle(bundleCheck.bundle, readFixture('command-result.json'))
    expect(commandResultCheck.status).toBe('valid')

    expect(checkWorkbookRunResultDescription(readFixture('strict-run-success.json')).status).toBe('valid')
    expect(checkWorkbookRunResultDescription(readFixture('strict-run-failure.json')).status).toBe('valid')

    const readbackProofCheck = checkWorkbookReadbackProof(readFixture('readback-proof.json'))
    expect(readbackProofCheck.status).toBe('valid')
    if (readbackProofCheck.status !== 'valid') {
      throw new Error('readback-proof fixture failed validation')
    }
    expect(readbackProofCheck.proof.checks).toEqual([
      expect.objectContaining({
        status: 'passed',
        kind: 'formulaEquals',
        proof: expect.objectContaining({
          source: 'readback',
          formula: 'input*factor',
        }),
      }),
    ])
  })
})
