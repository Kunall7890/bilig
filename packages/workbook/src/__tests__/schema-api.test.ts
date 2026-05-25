import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  checkPlanData,
  checkWorkbookCommandBundle,
  checkWorkbookCommandResultForBundle,
  checkWorkbookRunResultDescription,
  verifyWorkbookReadbacks,
  workbookJsonSchemaBundleHash,
  workbookJsonSchemaHash,
  workbookJsonSchemaHashes,
  workbookJsonSchemaNames,
  workbookJsonSchemas,
  workbookJsonSchemaVersion,
  workbookPlanId,
  type WorkbookCheckResult,
  type WorkbookRunReadback,
} from '../index.js'

const fixtureRoot = new URL('../../fixtures/', import.meta.url)

function readFixture(name: string): unknown {
  return JSON.parse(readFileSync(new URL(name, fixtureRoot), 'utf8'))
}

function requiredArray(value: unknown, path: string): readonly unknown[] {
  if (!Array.isArray(value)) {
    throw new Error(`${path} must be an array`)
  }
  return value
}

function requiredRecord(value: unknown, path: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${path} must be an object`)
  }
  return value
}

function requiredArrayOf<T>(value: unknown, path: string, isEntry: (entry: unknown) => entry is T): readonly T[] {
  return requiredArray(value, path).map((entry, index) => {
    if (!isEntry(entry)) {
      throw new Error(`${path}[${index}] has the wrong fixture shape`)
    }
    return entry
  })
}

function ownDataValue(value: object, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key)
  if (descriptor === undefined || !('value' in descriptor)) {
    return undefined
  }
  return descriptor.value
}

function isCheckFixture(value: unknown): value is WorkbookCheckResult {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  const status = ownDataValue(value, 'status')
  return (
    (status === 'planned' || status === 'passed' || status === 'failed') &&
    typeof ownDataValue(value, 'kind') === 'string' &&
    typeof ownDataValue(value, 'message') === 'string'
  )
}

function isReadbackFixture(value: unknown): value is WorkbookRunReadback {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return false
  }
  return Object.hasOwn(value, 'target')
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

    const readbackProof = requiredRecord(readFixture('readback-proof.json'), 'readback-proof')
    const checks = requiredArrayOf(readbackProof['checks'], 'readback-proof.checks', isCheckFixture)
    const readbacks = requiredArrayOf(readbackProof['readbacks'], 'readback-proof.readbacks', isReadbackFixture)
    const readbackVerification = verifyWorkbookReadbacks(checks, readbacks)
    expect(readbackVerification.status).toBe('passed')
    expect(readbackVerification.checks).toEqual([
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
