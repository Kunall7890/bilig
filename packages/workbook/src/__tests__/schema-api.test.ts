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
} from '../index.js'

const fixtureRoot = new URL('../../fixtures/', import.meta.url)

function readFixture(name: string): unknown {
  return JSON.parse(readFileSync(new URL(name, fixtureRoot), 'utf8'))
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
