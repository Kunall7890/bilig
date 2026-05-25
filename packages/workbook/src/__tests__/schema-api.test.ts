import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  checkPlanData,
  checkRuntimeRequirements,
  checkWorkbookCommandBundle,
  checkWorkbookCommandResultForBundle,
  checkWorkbookModelDescription,
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
const sourceRoot = new URL('../', import.meta.url)

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

function arrayEntry(value: unknown, key: string): readonly unknown[] {
  if (!isRecord(value)) {
    throw new Error(`expected object before reading ${key}`)
  }
  const entry = value[key]
  if (!Array.isArray(entry)) {
    throw new Error(`expected ${key} to be an array`)
  }
  return entry
}

function constKindFromSchema(value: unknown): string {
  const properties = objectEntry(value, 'properties')
  const kind = objectEntry(properties, 'kind')
  if (typeof kind['const'] !== 'string') {
    throw new Error('expected schema variant to have a string kind const')
  }
  return kind['const']
}

function workbookOpKindsFromSource(): readonly string[] {
  const source = readFileSync(new URL('ops.ts', sourceRoot), 'utf8')
  return [...source.matchAll(/\bkind:\s*'([^']+)'/gu)].map((match) => {
    const kind = match[1]
    if (kind === undefined) {
      throw new Error('expected WorkbookOp kind match')
    }
    return kind
  })
}

describe('@bilig/workbook schema api', () => {
  it('exports frozen JSON schema artifacts with deterministic hashes', () => {
    expect(workbookJsonSchemaVersion).toBe('bilig-workbook-json-schema-v1')
    expect(workbookJsonSchemaNames).toEqual([
      'refData',
      'modelDescription',
      'planData',
      'runtimeRequirements',
      'commandBundle',
      'commandResult',
      'runResult',
      'readbackProof',
    ])
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

  it('publishes engine ops as a known discriminated union', () => {
    const defs = objectEntry(workbookJsonSchemas.planData, '$defs')
    const engineOp = objectEntry(defs, 'engineOp')
    const variants = arrayEntry(engineOp, 'oneOf')
    const kinds = variants.map(constKindFromSchema)

    expect(kinds).toEqual(workbookOpKindsFromSource())

    const setCellValue = variants.find((entry) => constKindFromSchema(entry) === 'setCellValue')
    expect(setCellValue).toMatchObject({
      type: 'object',
      required: ['kind', 'sheetName', 'address', 'value'],
      properties: {
        value: { $ref: '#/$defs/literalInput' },
      },
    })
    const unknownKindSchema = variants.find((entry) => constKindFromSchema(entry) === 'makeRevenueModel')
    expect(unknownKindSchema).toBeUndefined()
  })

  it('publishes transported checks and format commands with checker-shaped payload schemas', () => {
    const defs = objectEntry(workbookJsonSchemas.planData, '$defs')
    expect(objectEntry(defs, 'literalInput')).toEqual({
      oneOf: [{ type: 'null' }, { type: 'boolean' }, { type: 'number' }, { type: 'string' }],
    })

    const checkExpectation = objectEntry(defs, 'checkExpectation')
    expect(checkExpectation).toMatchObject({
      oneOf: [
        {
          required: ['kind', 'value'],
          additionalProperties: false,
          properties: {
            kind: { const: 'valueEquals' },
            value: { $ref: '#/$defs/literalInput' },
          },
        },
        {
          required: ['kind', 'formula', 'inputs', 'labels'],
          additionalProperties: false,
          properties: {
            kind: { const: 'formulaEquals' },
            inputs: { type: 'array', items: { $ref: '#/$defs/refData' } },
            labels: { type: 'array', items: { $ref: '#/$defs/formulaLabel' } },
          },
        },
      ],
    })

    const command = objectEntry(defs, 'command')
    const commandVariants = arrayEntry(command, 'oneOf')
    const writeValue = commandVariants.find((entry) => constKindFromSchema(entry) === 'writeValue')
    const format = commandVariants.find((entry) => constKindFromSchema(entry) === 'format')
    expect(writeValue).toMatchObject({
      properties: {
        value: { $ref: '#/$defs/literalInput' },
      },
    })
    expect(format).toMatchObject({
      properties: {
        style: { $ref: '#/$defs/cellStylePatch' },
      },
    })

    const stylePatch = objectEntry(defs, 'cellStylePatch')
    const styleProperties = objectEntry(stylePatch, 'properties')
    const fontOneOf = arrayEntry(objectEntry(styleProperties, 'font'), 'oneOf')
    const fontObject = fontOneOf.find((entry) => isRecord(entry) && entry['type'] === 'object')
    expect(fontObject).toMatchObject({
      properties: {
        bold: { oneOf: [{ type: 'boolean' }, { type: 'null' }] },
      },
    })
  })

  it('publishes model manifest and action input description schemas for agent tool discovery', () => {
    const schema = workbookJsonSchemas.modelDescription
    expect(schema).toMatchObject({
      type: 'object',
      required: ['name', 'actions', 'actionDetails', 'hasChecks'],
      additionalProperties: false,
    })
    const properties = objectEntry(schema, 'properties')
    expect(properties['actionDetails']).toEqual({
      type: 'array',
      items: { $ref: '#/$defs/actionInspection' },
    })

    const defs = objectEntry(schema, '$defs')
    const actionInspection = objectEntry(defs, 'actionInspection')
    const actionInspectionProperties = objectEntry(actionInspection, 'properties')
    expect(actionInspectionProperties['input']).toEqual({ $ref: '#/$defs/actionInputDescription' })

    const actionInputDescription = objectEntry(defs, 'actionInputDescription')
    const actionInputProperties = objectEntry(actionInputDescription, 'properties')
    expect(actionInputProperties['kind']).toEqual({ enum: ['json', 'object', 'array', 'string', 'number', 'boolean', 'null'] })
    expect(actionInputProperties['values']).toEqual({
      type: 'array',
      minItems: 1,
      items: { $ref: '#/$defs/actionInput' },
    })
    expect(actionInputProperties['additionalProperties']).toEqual({ type: 'boolean' })
    expect(actionInputProperties['examples']).toEqual({
      type: 'array',
      minItems: 1,
      items: { $ref: '#/$defs/actionInput' },
    })
  })

  it('publishes the runtime requirements schema as an adapter handoff contract', () => {
    const schema = workbookJsonSchemas.runtimeRequirements
    expect(schema).toMatchObject({
      type: 'object',
      required: ['modelName', 'actionName', 'requirements'],
      additionalProperties: false,
    })
    const properties = objectEntry(schema, 'properties')
    expect(properties['requirements']).toEqual({
      type: 'array',
      items: { $ref: '#/$defs/runtimeRequirement' },
    })

    const requirement = objectEntry(objectEntry(schema, '$defs'), 'runtimeRequirement')
    expect(requirement).toMatchObject({
      type: 'object',
      required: ['kind', 'capability', 'message'],
      additionalProperties: false,
    })
    const requirementProperties = objectEntry(requirement, 'properties')
    expect(requirementProperties['kind']).toEqual({ enum: ['apply', 'read', 'verify'] })
    expect(requirementProperties['capability']).toEqual({
      enum: ['writeFormula', 'writeValue', 'format', 'clear', 'applyOp', 'read', 'verifyCheck'],
    })
    expect(requirementProperties['target']).toEqual({ $ref: '#/$defs/refData' })
    expect(requirementProperties['refs']).toEqual({
      type: 'array',
      items: { $ref: '#/$defs/refData' },
    })
  })

  it('keeps checked-in contract fixtures aligned with public validators', () => {
    const validPlan = readFixture('valid-plan.json')
    const validPlanCheck = checkPlanData(validPlan)
    expect(validPlanCheck.status).toBe('valid')
    if (validPlanCheck.status !== 'valid') {
      throw new Error('valid-plan fixture failed validation')
    }
    expect(workbookPlanId(validPlanCheck.plan)).toMatch(/^bilig-plan-v1:[0-9a-f]{32}$/u)

    const modelDescriptionCheck = checkWorkbookModelDescription(readFixture('model-description.json'))
    expect(modelDescriptionCheck.status).toBe('valid')

    const invalidPlanCheck = checkPlanData(readFixture('invalid-plan.json'))
    expect(invalidPlanCheck.status).toBe('invalid')

    const requirementsCheck = checkRuntimeRequirements(readFixture('runtime-requirements.json'))
    expect(requirementsCheck.status).toBe('valid')

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
