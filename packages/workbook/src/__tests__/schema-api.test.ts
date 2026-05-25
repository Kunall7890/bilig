import { readFileSync } from 'node:fs'
import Ajv2020 from 'ajv/dist/2020.js'
import { describe, expect, it } from 'vitest'
import {
  checkPlanData,
  checkRuntimeRequirements,
  checkWorkbookCommandBundle,
  checkWorkbookCommandResult,
  checkWorkbookCommandResultForBundle,
  checkWorkbookRefData,
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
const nonNegativeSafeIntegerSchema = Object.freeze({
  type: 'integer',
  minimum: 0,
  maximum: Number.MAX_SAFE_INTEGER,
})
const exactStringSchema = Object.freeze({
  type: 'string',
  minLength: 1,
  pattern: '^(?!\\s)(?![\\s\\S]*\\s$)[\\s\\S]+$',
})

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

function workbookOpVariant(variants: readonly unknown[], kind: string): Record<string, unknown> {
  const variant = variants.find((entry) => constKindFromSchema(entry) === kind)
  if (!isRecord(variant)) {
    throw new Error(`expected ${kind} schema variant`)
  }
  return variant
}

function schemaProperty(schema: unknown, key: string): unknown {
  return objectEntry(schema, 'properties')[key]
}

function validateWithJsonSchema(name: keyof typeof workbookJsonSchemas, payload: unknown): boolean {
  const ajv = new Ajv2020({ allErrors: true, strict: false })
  const validate = ajv.compile(workbookJsonSchemas[name])
  return validate(payload)
}

function expectInvalidRefDataParity(payload: unknown): void {
  expect(validateWithJsonSchema('refData', payload)).toBe(false)
  expect(checkWorkbookRefData(payload).status).toBe('invalid')
}

function expectInvalidCommandBundleParity(payload: unknown): void {
  expect(validateWithJsonSchema('commandBundle', payload)).toBe(false)
  expect(checkWorkbookCommandBundle(payload).status).toBe('invalid')
}

function expectInvalidCommandResultParity(payload: unknown): void {
  expect(validateWithJsonSchema('commandResult', payload)).toBe(false)
  expect(checkWorkbookCommandResult(payload).status).toBe('invalid')
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
    expect(applyProperties['baseRevision']).toEqual(nonNegativeSafeIntegerSchema)
    expect(applyProperties['revision']).toEqual(nonNegativeSafeIntegerSchema)

    expect(commandReceipt).toMatchObject({
      type: 'object',
      required: ['commandIndex', 'commandKind', 'commandDigest', 'previewOps', 'appliedOps'],
      additionalProperties: false,
    })
    expect(commandReceiptProperties['commandIndex']).toEqual(nonNegativeSafeIntegerSchema)
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
        id: exactStringSchema,
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

  it('publishes checker-shaped payload schemas for known engine ops', () => {
    const defs = objectEntry(workbookJsonSchemas.planData, '$defs')
    const variants = arrayEntry(objectEntry(defs, 'engineOp'), 'oneOf')
    const looseObject = { type: 'object', additionalProperties: true }

    const calculationSettings = objectEntry(schemaProperty(workbookOpVariant(variants, 'setCalculationSettings'), 'settings'), 'properties')
    expect(calculationSettings['mode']).toEqual({ enum: ['automatic', 'manual'] })
    expect(calculationSettings['compatibilityMode']).toEqual({ enum: ['excel-modern', 'odf-1.4'] })

    const volatileContext = objectEntry(schemaProperty(workbookOpVariant(variants, 'setVolatileContext'), 'context'), 'properties')
    expect(volatileContext['recalcEpoch']).toEqual({ type: 'number' })

    const directPayloads: ReadonlyArray<readonly [string, string]> = [
      ['setCalculationSettings', 'settings'],
      ['setVolatileContext', 'context'],
      ['setSheetProtection', 'protection'],
      ['setDataValidation', 'validation'],
      ['upsertConditionalFormat', 'format'],
      ['setConditionalFormatArtifacts', 'artifacts'],
      ['upsertRangeProtection', 'protection'],
      ['upsertCommentThread', 'thread'],
      ['upsertNote', 'note'],
      ['upsertHyperlink', 'hyperlink'],
      ['upsertDefinedName', 'value'],
      ['upsertTable', 'table'],
      ['upsertChart', 'chart'],
      ['upsertImage', 'image'],
      ['upsertShape', 'shape'],
    ]
    for (const [kind, key] of directPayloads) {
      expect(schemaProperty(workbookOpVariant(variants, kind), key), `${kind}.${key}`).not.toEqual(looseObject)
    }

    const pivotValues = objectEntry(schemaProperty(workbookOpVariant(variants, 'upsertPivotTable'), 'values'), 'items')
    expect(pivotValues).not.toEqual(looseObject)
    expect(objectEntry(pivotValues, 'properties')).toMatchObject({
      sourceColumn: { type: 'string', minLength: 1 },
      summarizeBy: { enum: ['sum', 'count'] },
    })
  })

  it('exports JSON schemas that external validators enforce safe integer bounds', () => {
    for (const name of workbookJsonSchemaNames) {
      expect(() => new Ajv2020({ allErrors: true, strict: false }).compile(workbookJsonSchemas[name])).not.toThrow()
    }

    expect(validateWithJsonSchema('commandBundle', readFixture('command-bundle.json'))).toBe(true)

    const unsafeInteger = Number.MAX_SAFE_INTEGER + 1
    const targetRange = Object.freeze({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A1',
    })
    const invalidBundles = [
      {
        targetRevision: unsafeInteger,
        idempotencyKey: 'unsafe-target-revision',
        commands: [
          {
            kind: 'op',
            destructive: true,
            op: { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 },
          },
        ],
      },
      {
        targetRevision: 1,
        idempotencyKey: 'unsafe-scope',
        scope: { maxTouchedCells: unsafeInteger },
        commands: [
          {
            kind: 'op',
            destructive: true,
            touchedRanges: [targetRange],
            op: { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
          },
        ],
      },
      {
        targetRevision: 1,
        idempotencyKey: 'unsafe-known-op',
        commands: [
          {
            kind: 'op',
            destructive: true,
            op: { kind: 'insertRows', sheetName: 'Sheet1', start: unsafeInteger, count: 1 },
          },
        ],
      },
    ] as const

    for (const bundle of invalidBundles) {
      expect(validateWithJsonSchema('commandBundle', bundle), JSON.stringify(bundle)).toBe(false)
    }
  })

  it('exports ref schemas that external validators enforce nested ref roles', () => {
    const table = {
      kind: 'table',
      id: 'table-sales',
      label: 'Sales',
      headers: ['Region', 'Amount'],
    } as const
    const rows = {
      kind: 'rows',
      id: 'rows-emea',
      label: 'EMEA rows',
      table,
      where: {
        column: 'Region',
        op: 'eq',
        value: 'EMEA',
      },
    } as const
    const range = {
      kind: 'range',
      id: 'range-a1',
      label: 'Sheet1!A1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    } as const
    const column = {
      kind: 'column',
      id: 'column-amount',
      label: 'Sales.Amount',
      table,
      rows,
      name: 'Amount',
    } as const

    expect(validateWithJsonSchema('refData', column)).toBe(true)
    expect(checkWorkbookRefData(column).status).toBe('valid')

    const invalidColumnTable = { ...column, table: range }
    const invalidColumnRows = { ...column, rows: table }
    const invalidRowsTable = { ...rows, table: range }

    for (const ref of [invalidColumnTable, invalidColumnRows, invalidRowsTable]) {
      expect(validateWithJsonSchema('refData', ref), JSON.stringify(ref)).toBe(false)
      expect(checkWorkbookRefData(ref).status).toBe('invalid')
    }
  })

  it('exports plan schemas that reject empty format intent like the checker', () => {
    const target = {
      kind: 'range',
      id: 'range-a1',
      label: 'Sheet1!A1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    } as const
    const basePlan = {
      modelName: 'format-schema-model',
      actionName: 'format',
      refsUsed: [target],
      ops: [],
      changed: [],
      checks: [],
    } as const
    const invalidPlans = [
      {
        ...basePlan,
        commands: [{ kind: 'format', target }],
      },
      {
        ...basePlan,
        commands: [{ kind: 'format', target, style: {} }],
      },
      {
        ...basePlan,
        commands: [{ kind: 'format', target, style: { font: {} } }],
      },
      {
        ...basePlan,
        commands: [{ kind: 'format', target, style: { borders: { top: {} } } }],
      },
    ] as const

    for (const plan of invalidPlans) {
      expect(validateWithJsonSchema('planData', plan), JSON.stringify(plan)).toBe(false)
      expect(checkPlanData(plan).status).toBe('invalid')
    }

    const validPlan = {
      ...basePlan,
      commands: [{ kind: 'format', target, style: { font: { bold: true } } }],
    } as const
    expect(validateWithJsonSchema('planData', validPlan)).toBe(true)
    expect(checkPlanData(validPlan).status).toBe('valid')
  })

  it('exports command-result schemas that reject loose undo and accepted settled fields', () => {
    const accepted = {
      status: 'accepted',
      targetRevision: 7,
      idempotencyKey: 'result-schema',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
    } as const
    expect(validateWithJsonSchema('commandResult', accepted)).toBe(true)

    const acceptedWithUndo = {
      ...accepted,
      undo: {
        id: 'undo-1',
      },
    } as const
    expect(validateWithJsonSchema('commandResult', acceptedWithUndo)).toBe(false)
    expect(checkWorkbookCommandResult(acceptedWithUndo).status).toBe('invalid')

    const appliedWithLooseRootUndo = {
      ...accepted,
      status: 'applied',
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          changedRanges: [],
        },
      ],
      matched: null,
      changedRanges: [],
      undo: {},
    } as const
    expect(validateWithJsonSchema('commandResult', appliedWithLooseRootUndo)).toBe(false)
    expect(checkWorkbookCommandResult(appliedWithLooseRootUndo).status).toBe('invalid')

    const appliedWithLooseReceiptUndo = {
      ...accepted,
      status: 'applied',
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
          undo: {},
        },
      ],
      matched: null,
      changedRanges: [],
    } as const
    expect(validateWithJsonSchema('commandResult', appliedWithLooseReceiptUndo)).toBe(false)
    expect(checkWorkbookCommandResult(appliedWithLooseReceiptUndo).status).toBe('invalid')
  })

  it('keeps command transport string schemas aligned with exact-string checks', () => {
    expectInvalidCommandBundleParity({
      targetRevision: 1,
      idempotencyKey: ' bundle ',
      commands: [
        {
          kind: 'op',
          destructive: true,
          op: { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
        },
      ],
    })
    expectInvalidCommandBundleParity({
      targetRevision: 1,
      idempotencyKey: 'bundle',
      commands: [
        {
          kind: 'request',
          id: ' command ',
          request: {
            featureId: 'cells',
            commandId: 'set-value',
          },
        },
      ],
    })
    expectInvalidCommandBundleParity({
      targetRevision: 1,
      idempotencyKey: 'bundle',
      commands: [
        {
          kind: 'request',
          request: {
            featureId: ' cells',
            commandId: 'set-value',
          },
        },
      ],
    })

    expectInvalidCommandResultParity({
      status: 'accepted',
      targetRevision: 1,
      idempotencyKey: ' result ',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
    })
    expectInvalidCommandResultParity({
      status: 'applied',
      targetRevision: 1,
      idempotencyKey: 'result',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: ' set-value',
          category: 'mutation',
          changedRanges: [],
        },
      ],
      matched: null,
      changedRanges: [],
    })
    expectInvalidCommandResultParity({
      status: 'rejected',
      targetRevision: 1,
      idempotencyKey: 'result',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
      receipts: [
        {
          status: 'rejected',
          featureId: 'cells',
          commandId: 'set-value',
          category: 'mutation',
          errors: [' failed'],
        },
      ],
      matched: null,
      changedRanges: [],
      errors: [' failed'],
    })
  })

  it('publishes command-bundle style record schemas that match op guards', () => {
    const defs = objectEntry(workbookJsonSchemas.commandBundle, '$defs')
    const variants = arrayEntry(objectEntry(defs, 'engineOp'), 'oneOf')
    const upsertCellStyle = workbookOpVariant(variants, 'upsertCellStyle')
    const style = schemaProperty(upsertCellStyle, 'style')
    const fill = schemaProperty(style, 'fill')

    expect(fill).toMatchObject({
      type: 'object',
      required: ['backgroundColor'],
      additionalProperties: true,
      properties: {
        backgroundColor: { type: 'string' },
      },
    })
    expect(
      checkWorkbookCommandBundle({
        targetRevision: 1,
        idempotencyKey: 'style-schema-parity',
        commands: [
          {
            kind: 'op',
            destructive: true,
            op: { kind: 'upsertCellStyle', style: { id: 'style-1', fill: {} } },
          },
        ],
      }),
    ).toMatchObject({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_command',
          path: 'commands[0].op',
        },
      ],
    })
  })

  it('keeps row-ref JSON Schema constraints aligned with ref-data checks', () => {
    const validRows = {
      kind: 'rows',
      id: 'rows-ready',
      label: 'Ready rows',
      where: { column: 'Status', op: 'eq', value: 'ready' },
    } as const
    expect(validateWithJsonSchema('refData', validRows)).toBe(true)
    expect(checkWorkbookRefData(validRows).status).toBe('valid')

    expectInvalidRefDataParity({
      ...validRows,
      where: { column: 'Status', op: 'contains', value: 123 },
    })
    expectInvalidRefDataParity({
      ...validRows,
      where: { column: 'Status', op: 'startsWith', value: ['ready'] },
    })
    expectInvalidRefDataParity({
      ...validRows,
      where: { column: 'Status', op: 'gt', value: false },
    })
    expectInvalidRefDataParity({
      ...validRows,
      where: { column: 'Status', op: 'lte', value: { amount: 10 } },
    })
  })

  it('keeps low-level op command schemas aligned with destructive checks', () => {
    const opCommand = {
      targetRevision: 1,
      idempotencyKey: 'op-schema-parity',
      commands: [
        {
          kind: 'op',
          op: { kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' },
        },
      ],
    } as const

    expectInvalidCommandBundleParity(opCommand)
    expect(checkWorkbookCommandBundle(opCommand)).toMatchObject({
      status: 'invalid',
      issues: [
        {
          code: 'destructive_not_confirmed',
          path: 'commands[0].destructive',
        },
      ],
    })

    const confirmed = {
      ...opCommand,
      commands: [
        {
          ...opCommand.commands[0],
          destructive: true,
        },
      ],
    } as const
    expect(validateWithJsonSchema('commandBundle', confirmed)).toBe(true)
    expect(checkWorkbookCommandBundle(confirmed).status).toBe('valid')
  })

  it('publishes comparison rule schemas with checker-matching value counts', () => {
    const defs = objectEntry(workbookJsonSchemas.planData, '$defs')
    const variants = arrayEntry(objectEntry(defs, 'engineOp'), 'oneOf')
    const setDataValidation = workbookOpVariant(variants, 'setDataValidation')
    const validation = schemaProperty(setDataValidation, 'validation')
    const validationRule = schemaProperty(validation, 'rule')
    const validationRuleVariants = arrayEntry(validationRule, 'oneOf')
    const validationComparisonRules = validationRuleVariants.filter((entry) => {
      const kind = schemaProperty(entry, 'kind')
      return isRecord(kind) && Array.isArray(kind['enum']) && kind['enum'].includes('whole')
    })

    expect(validationComparisonRules).toEqual([
      expect.objectContaining({
        properties: expect.objectContaining({
          operator: { enum: ['between', 'notBetween'] },
          values: {
            type: 'array',
            minItems: 2,
            maxItems: 2,
            items: { $ref: '#/$defs/literalInput' },
          },
        }),
      }),
      expect.objectContaining({
        properties: expect.objectContaining({
          operator: {
            enum: ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'],
          },
          values: {
            type: 'array',
            minItems: 1,
            maxItems: 1,
            items: { $ref: '#/$defs/literalInput' },
          },
        }),
      }),
    ])

    const upsertConditionalFormat = workbookOpVariant(variants, 'upsertConditionalFormat')
    const conditionalFormat = schemaProperty(upsertConditionalFormat, 'format')
    const conditionalFormatRule = schemaProperty(conditionalFormat, 'rule')
    const conditionalFormatRuleVariants = arrayEntry(conditionalFormatRule, 'oneOf')
    const cellIsRules = conditionalFormatRuleVariants.filter((entry) => {
      const kind = schemaProperty(entry, 'kind')
      return isRecord(kind) && kind['const'] === 'cellIs'
    })

    expect(cellIsRules).toEqual([
      expect.objectContaining({
        properties: expect.objectContaining({
          operator: { enum: ['between', 'notBetween'] },
          values: {
            type: 'array',
            minItems: 2,
            maxItems: 2,
            items: { $ref: '#/$defs/literalInput' },
          },
        }),
      }),
      expect.objectContaining({
        properties: expect.objectContaining({
          operator: {
            enum: ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'],
          },
          values: {
            type: 'array',
            minItems: 1,
            maxItems: 1,
            items: { $ref: '#/$defs/literalInput' },
          },
        }),
      }),
    ])
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
      anyOf: [{ required: ['style'] }, { required: ['numberFormat'] }],
      properties: {
        style: { $ref: '#/$defs/actionCellStylePatch' },
      },
    })

    const actionStylePatch = objectEntry(defs, 'actionCellStylePatch')
    expect(actionStylePatch).toMatchObject({
      allOf: [
        { $ref: '#/$defs/cellStylePatch' },
        {
          anyOf: expect.any(Array),
        },
      ],
    })

    const stylePatch = objectEntry(defs, 'cellStylePatch')
    const styleProperties = objectEntry(stylePatch, 'properties')
    const fontOneOf = arrayEntry(objectEntry(styleProperties, 'font'), 'oneOf')
    const fontObject = fontOneOf.find((entry) => isRecord(entry) && entry['type'] === 'object')
    expect(stylePatch).toMatchObject({
      additionalProperties: false,
    })
    expect(fontObject).toMatchObject({
      additionalProperties: false,
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
