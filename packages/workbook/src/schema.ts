import { workbookRunErrorCodes } from './result.js'
import { workbookActionInputDescriptionKinds } from './input.js'
import { cellStylePatchSchema, literalInputSchema, workbookOpSchema } from './op-schema.js'
import {
  workbookActionCellStylePatchSchema,
  workbookCommandReceiptSchema,
  workbookRefDataDefSchemas,
  workbookRefDataSchemaRefs,
  workbookUndoRefSchema,
} from './schema-fragments.js'

export type WorkbookJsonSchemaScalar = string | number | boolean | null
export type WorkbookJsonSchemaValue =
  | WorkbookJsonSchemaScalar
  | readonly WorkbookJsonSchemaValue[]
  | {
      readonly [key: string]: WorkbookJsonSchemaValue
    }

export type WorkbookJsonSchema = {
  readonly [key: string]: WorkbookJsonSchemaValue
}

export type WorkbookJsonSchemaName =
  | 'refData'
  | 'modelDescription'
  | 'planData'
  | 'runtimeRequirements'
  | 'commandBundle'
  | 'commandResult'
  | 'runResult'
  | 'readbackProof'

export const workbookJsonSchemaVersion = 'bilig-workbook-json-schema-v1'

export const workbookJsonSchemaNames = Object.freeze([
  'refData',
  'modelDescription',
  'planData',
  'runtimeRequirements',
  'commandBundle',
  'commandResult',
  'runResult',
  'readbackProof',
] as const satisfies readonly WorkbookJsonSchemaName[])

const jsonValue = Object.freeze({ $ref: '#/$defs/jsonValue' })
const literalInput = Object.freeze({ $ref: '#/$defs/literalInput' })
const refData = Object.freeze({ $ref: '#/$defs/refData' })
const cellRange = Object.freeze({ $ref: '#/$defs/cellRange' })
const engineOp = Object.freeze({ $ref: '#/$defs/engineOp' })
const actionInput = Object.freeze({ $ref: '#/$defs/actionInput' })
const actionInputDescription = Object.freeze({ $ref: '#/$defs/actionInputDescription' })
const exactString = Object.freeze({
  type: 'string',
  minLength: 1,
  pattern: '^(?!\\s)(?![\\s\\S]*\\s$)[\\s\\S]+$',
} as const)
const nonNegativeSafeInteger = Object.freeze({
  type: 'integer',
  minimum: 0,
  maximum: Number.MAX_SAFE_INTEGER,
} as const)

function defs(extra: Record<string, WorkbookJsonSchemaValue> = {}): WorkbookJsonSchemaValue {
  return {
    jsonValue: {
      oneOf: [
        { type: 'null' },
        { type: 'boolean' },
        { type: 'number' },
        { type: 'string' },
        { type: 'array', items: { $ref: '#/$defs/jsonValue' } },
        { type: 'object', additionalProperties: { $ref: '#/$defs/jsonValue' } },
      ],
    },
    literalInput: literalInputSchema,
    actionInput: { $ref: '#/$defs/jsonValue' },
    actionInputDescription: {
      type: 'object',
      required: ['kind'],
      additionalProperties: false,
      properties: {
        kind: { enum: workbookActionInputDescriptionKinds },
        description: { type: 'string', minLength: 1 },
        required: { type: 'boolean' },
        fields: {
          type: 'object',
          propertyNames: { minLength: 1 },
          additionalProperties: { $ref: '#/$defs/actionInputDescription' },
        },
        items: { $ref: '#/$defs/actionInputDescription' },
        values: { type: 'array', minItems: 1, items: actionInput },
        min: { type: 'number' },
        max: { type: 'number' },
        minLength: nonNegativeSafeInteger,
        maxLength: nonNegativeSafeInteger,
        pattern: { type: 'string' },
        minItems: nonNegativeSafeInteger,
        maxItems: nonNegativeSafeInteger,
        additionalProperties: { type: 'boolean' },
        default: actionInput,
        examples: { type: 'array', minItems: 1, items: actionInput },
      },
      allOf: [
        {
          if: { properties: { kind: { const: 'object' } }, required: ['kind'] },
          // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
          then: true,
          else: { not: { anyOf: [{ required: ['fields'] }, { required: ['additionalProperties'] }] } },
        },
        {
          if: { properties: { kind: { const: 'array' } }, required: ['kind'] },
          // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
          then: true,
          else: { not: { anyOf: [{ required: ['items'] }, { required: ['minItems'] }, { required: ['maxItems'] }] } },
        },
        {
          if: { properties: { kind: { const: 'number' } }, required: ['kind'] },
          // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
          then: true,
          else: { not: { anyOf: [{ required: ['min'] }, { required: ['max'] }] } },
        },
        {
          if: { properties: { kind: { const: 'string' } }, required: ['kind'] },
          // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
          then: true,
          else: { not: { anyOf: [{ required: ['minLength'] }, { required: ['maxLength'] }, { required: ['pattern'] }] } },
        },
      ],
    },
    cellRange: {
      type: 'object',
      required: ['sheetName', 'startAddress', 'endAddress'],
      additionalProperties: false,
      properties: {
        sheetName: { type: 'string', minLength: 1 },
        startAddress: { type: 'string', minLength: 1 },
        endAddress: { type: 'string', minLength: 1 },
      },
    },
    cellStylePatch: cellStylePatchSchema,
    actionCellStylePatch: workbookActionCellStylePatchSchema,
    formulaLabel: {
      type: 'object',
      required: ['name', 'ref'],
      additionalProperties: false,
      properties: {
        name: { type: 'string', minLength: 1 },
        ref: refData,
      },
    },
    engineOp: workbookOpSchema,
    ...workbookRefDataDefSchemas,
    refData: {
      oneOf: workbookRefDataSchemaRefs,
    },
    concreteRefData: {
      type: 'object',
      required: ['kind', 'id', 'label', 'range'],
      additionalProperties: false,
      properties: {
        kind: { const: 'range' },
        id: { type: 'string', minLength: 1 },
        label: { type: 'string', minLength: 1 },
        range: { $ref: '#/$defs/cellRange' },
      },
    },
    resolvedRefValue: {
      oneOf: [{ $ref: '#/$defs/concreteRefData' }, { type: 'array', items: { $ref: '#/$defs/concreteRefData' } }],
    },
    commandResolvedRefs: {
      type: 'object',
      additionalProperties: false,
      properties: {
        target: { $ref: '#/$defs/resolvedRefValue' },
        inputs: { type: 'array', items: { $ref: '#/$defs/resolvedRefValue' } },
      },
    },
    ...extra,
  }
}

function schema(name: WorkbookJsonSchemaName, value: WorkbookJsonSchema): WorkbookJsonSchema {
  return deepFreeze({
    $schema: 'https://json-schema.org/draft/2020-12/schema',
    $id: `https://bilig.proompteng.ai/schemas/workbook/${workbookJsonSchemaVersion}/${name}.schema.json`,
    title: `@bilig/workbook ${name}`,
    ...value,
  })
}

const checkDataSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['status', 'kind', 'message'],
  properties: {
    status: { enum: ['planned', 'passed', 'failed'] },
    kind: { type: 'string', minLength: 1 },
    target: refData,
    refs: { type: 'array', items: refData },
    message: { type: 'string', minLength: 1 },
    expectation: { $ref: '#/$defs/checkExpectation' },
    proof: actionInput,
  },
  additionalProperties: false,
}

const checkExpectationSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    {
      type: 'object',
      required: ['kind', 'value'],
      additionalProperties: false,
      properties: {
        kind: { const: 'valueEquals' },
        value: literalInput,
      },
    },
    {
      type: 'object',
      required: ['kind', 'formula', 'inputs', 'labels'],
      additionalProperties: false,
      properties: {
        kind: { const: 'formulaEquals' },
        formula: { type: 'string' },
        inputs: { type: 'array', items: refData },
        labels: { type: 'array', items: { $ref: '#/$defs/formulaLabel' } },
      },
    },
  ],
}

const changeDataSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['kind', 'message'],
  additionalProperties: false,
  properties: {
    kind: { type: 'string', minLength: 1 },
    target: refData,
    message: { type: 'string', minLength: 1 },
  },
}

const unverifiedDataSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['kind', 'message'],
  additionalProperties: false,
  properties: {
    kind: { enum: ['apply', 'plan'] },
    message: { type: 'string', minLength: 1 },
  },
}

const runtimeRequirementSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['kind', 'capability', 'message'],
  additionalProperties: false,
  properties: {
    kind: { enum: ['apply', 'read', 'verify'] },
    capability: { enum: ['writeFormula', 'writeValue', 'format', 'clear', 'applyOp', 'read', 'verifyCheck'] },
    commandIndex: nonNegativeSafeInteger,
    checkIndex: nonNegativeSafeInteger,
    opIndex: nonNegativeSafeInteger,
    opKind: { type: 'string', minLength: 1 },
    checkKind: { type: 'string', minLength: 1 },
    target: refData,
    refs: { type: 'array', items: refData },
    message: { type: 'string', minLength: 1 },
  },
  allOf: [
    {
      if: { properties: { kind: { const: 'apply' } }, required: ['kind'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: { properties: { capability: { enum: ['writeFormula', 'writeValue', 'format', 'clear', 'applyOp'] } } },
    },
    {
      if: { properties: { kind: { const: 'read' } }, required: ['kind'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: { properties: { capability: { const: 'read' } } },
    },
    {
      if: { properties: { kind: { const: 'verify' } }, required: ['kind'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: { properties: { capability: { const: 'verifyCheck' } } },
    },
  ],
}

const applySummarySchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['matched'],
  additionalProperties: false,
  properties: {
    matched: { oneOf: [{ type: 'boolean' }, { type: 'null' }] },
    planId: { type: 'string', minLength: 1 },
    baseRevision: nonNegativeSafeInteger,
    revision: nonNegativeSafeInteger,
    previewOps: { type: 'array', items: engineOp },
    appliedOps: { type: 'array', items: engineOp },
    commandReceipts: { type: 'array', items: { $ref: '#/$defs/applyCommandReceipt' } },
    proof: actionInput,
  },
}

const applyCommandReceiptSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['commandIndex', 'commandKind', 'commandDigest', 'previewOps', 'appliedOps'],
  additionalProperties: false,
  properties: {
    commandIndex: nonNegativeSafeInteger,
    commandKind: { type: 'string', minLength: 1 },
    commandDigest: { type: 'string', minLength: 1 },
    previewOps: { type: 'array', items: engineOp },
    appliedOps: { type: 'array', items: engineOp },
    noop: {
      type: 'object',
      required: ['reason', 'proof'],
      additionalProperties: false,
      properties: {
        reason: { enum: ['already_satisfied'] },
        message: { type: 'string', minLength: 1 },
        proof: {
          type: 'object',
          required: ['source', 'evidence', 'opCount', 'commandKind', 'commandDigest', 'effect'],
          additionalProperties: true,
          properties: {
            source: { type: 'string', minLength: 1 },
            evidence: { type: 'string', minLength: 1 },
            opCount: { const: 0 },
            commandKind: { type: 'string', minLength: 1 },
            commandDigest: { type: 'string', minLength: 1 },
            effect: actionInput,
          },
        },
      },
    },
    resolvedRefs: { $ref: '#/$defs/commandResolvedRefs' },
    formulaLabels: {
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
    },
    proof: actionInput,
  },
  allOf: [
    {
      if: { required: ['noop'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: {
        properties: {
          previewOps: { type: 'array', maxItems: 0, items: engineOp },
          appliedOps: { type: 'array', maxItems: 0, items: engineOp },
        },
      },
    },
  ],
}

export const workbookJsonSchemas = deepFreeze({
  refData: schema('refData', {
    $defs: defs(),
    ...refData,
  }),

  modelDescription: schema('modelDescription', {
    $defs: defs({
      actionInspection: {
        type: 'object',
        required: ['name'],
        additionalProperties: false,
        properties: {
          name: { type: 'string', minLength: 1 },
          description: { type: 'string', minLength: 1 },
          input: actionInputDescription,
        },
      },
    }),
    type: 'object',
    required: ['name', 'actions', 'actionDetails', 'hasChecks'],
    additionalProperties: false,
    properties: {
      name: { type: 'string', minLength: 1 },
      description: { type: 'string', minLength: 1 },
      actions: { type: 'array', items: { type: 'string', minLength: 1 } },
      actionDetails: { type: 'array', items: { $ref: '#/$defs/actionInspection' } },
      hasChecks: { type: 'boolean' },
    },
  }),

  planData: schema('planData', {
    $defs: defs({
      checkExpectation: checkExpectationSchema,
      check: checkDataSchema,
      command: {
        oneOf: [
          {
            type: 'object',
            required: ['kind', 'target', 'formula', 'inputs', 'labels'],
            additionalProperties: false,
            properties: {
              kind: { const: 'writeFormula' },
              target: refData,
              formula: { type: 'string' },
              inputs: { type: 'array', items: refData },
              labels: { type: 'array', items: { $ref: '#/$defs/formulaLabel' } },
            },
          },
          {
            type: 'object',
            required: ['kind', 'target', 'value'],
            additionalProperties: false,
            properties: { kind: { const: 'writeValue' }, target: refData, value: literalInput },
          },
          {
            type: 'object',
            required: ['kind', 'target'],
            additionalProperties: false,
            properties: {
              kind: { const: 'format' },
              target: refData,
              style: { $ref: '#/$defs/actionCellStylePatch' },
              numberFormat: { oneOf: [{ type: 'string' }, { type: 'null' }] },
            },
            anyOf: [{ required: ['style'] }, { required: ['numberFormat'] }],
          },
          {
            type: 'object',
            required: ['kind', 'target'],
            additionalProperties: false,
            properties: { kind: { const: 'clear' }, target: refData },
          },
          {
            type: 'object',
            required: ['kind', 'op'],
            additionalProperties: false,
            properties: {
              kind: { const: 'op' },
              op: engineOp,
              target: refData,
              message: { type: 'string' },
            },
          },
        ],
      },
      change: changeDataSchema,
    }),
    type: 'object',
    required: ['modelName', 'actionName', 'refsUsed', 'commands', 'ops', 'changed', 'checks'],
    additionalProperties: false,
    properties: {
      modelName: { type: 'string', minLength: 1 },
      actionName: { type: 'string', minLength: 1 },
      input: actionInput,
      refsUsed: { type: 'array', items: refData },
      commands: { type: 'array', items: { $ref: '#/$defs/command' } },
      ops: { type: 'array', items: engineOp },
      changed: { type: 'array', items: { $ref: '#/$defs/change' } },
      checks: { type: 'array', items: { $ref: '#/$defs/check' } },
    },
  }),

  runtimeRequirements: schema('runtimeRequirements', {
    $defs: defs({
      runtimeRequirement: runtimeRequirementSchema,
    }),
    type: 'object',
    required: ['modelName', 'actionName', 'requirements'],
    additionalProperties: false,
    properties: {
      modelName: { type: 'string', minLength: 1 },
      actionName: { type: 'string', minLength: 1 },
      requirements: { type: 'array', items: { $ref: '#/$defs/runtimeRequirement' } },
    },
  }),

  commandBundle: schema('commandBundle', {
    $defs: defs({
      commandRequest: {
        type: 'object',
        required: ['featureId', 'commandId'],
        additionalProperties: false,
        properties: {
          featureId: exactString,
          commandId: exactString,
          category: { enum: ['command', 'operation', 'mutation'] },
          mode: { enum: ['preview', 'apply', 'applyAndVerify'] },
          input: actionInput,
        },
      },
      bundleCommand: {
        oneOf: [
          {
            type: 'object',
            required: ['kind', 'request'],
            additionalProperties: false,
            properties: {
              kind: { const: 'request' },
              id: exactString,
              touchedRanges: { type: 'array', items: cellRange },
              destructive: { type: 'boolean' },
              request: { $ref: '#/$defs/commandRequest' },
            },
            allOf: [
              {
                if: {
                  properties: {
                    request: {
                      type: 'object',
                      properties: { category: { const: 'mutation' } },
                      required: ['category'],
                    },
                  },
                  required: ['request'],
                },
                // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
                then: {
                  required: ['destructive'],
                  properties: { destructive: { const: true } },
                },
              },
              {
                if: {
                  properties: {
                    request: {
                      type: 'object',
                      properties: { mode: { enum: ['apply', 'applyAndVerify'] } },
                      required: ['mode'],
                    },
                  },
                  required: ['request'],
                },
                // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
                then: {
                  required: ['destructive'],
                  properties: { destructive: { const: true } },
                },
              },
            ],
          },
          {
            type: 'object',
            required: ['kind', 'destructive', 'op'],
            additionalProperties: false,
            properties: {
              kind: { const: 'op' },
              id: exactString,
              touchedRanges: { type: 'array', items: cellRange },
              destructive: { const: true },
              op: engineOp,
            },
          },
        ],
      },
    }),
    type: 'object',
    required: ['targetRevision', 'idempotencyKey', 'commands'],
    additionalProperties: false,
    properties: {
      id: exactString,
      targetRevision: nonNegativeSafeInteger,
      idempotencyKey: exactString,
      scope: {
        type: 'object',
        additionalProperties: false,
        properties: { maxTouchedCells: nonNegativeSafeInteger },
      },
      commands: { type: 'array', minItems: 1, items: { $ref: '#/$defs/bundleCommand' } },
    },
    allOf: [
      {
        if: {
          properties: {
            scope: {
              type: 'object',
              required: ['maxTouchedCells'],
            },
          },
          required: ['scope'],
        },
        // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
        then: {
          properties: {
            commands: {
              items: {
                allOf: [
                  {
                    if: {
                      properties: { kind: { const: 'op' } },
                      required: ['kind'],
                    },
                    // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
                    then: {
                      required: ['touchedRanges'],
                      properties: { touchedRanges: { type: 'array', minItems: 1, items: cellRange } },
                    },
                  },
                  {
                    if: {
                      properties: {
                        kind: { const: 'request' },
                        request: {
                          type: 'object',
                          properties: { category: { const: 'mutation' } },
                          required: ['category'],
                        },
                      },
                      required: ['kind', 'request'],
                    },
                    // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
                    then: {
                      required: ['touchedRanges'],
                      properties: { touchedRanges: { type: 'array', minItems: 1, items: cellRange } },
                    },
                  },
                  {
                    if: {
                      properties: {
                        kind: { const: 'request' },
                        request: {
                          type: 'object',
                          properties: { mode: { enum: ['apply', 'applyAndVerify'] } },
                          required: ['mode'],
                        },
                      },
                      required: ['kind', 'request'],
                    },
                    // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
                    then: {
                      required: ['touchedRanges'],
                      properties: { touchedRanges: { type: 'array', minItems: 1, items: cellRange } },
                    },
                  },
                ],
              },
            },
          },
        },
      },
    ],
  }),

  commandResult: schema('commandResult', {
    $defs: defs({
      undo: workbookUndoRefSchema,
      receipt: workbookCommandReceiptSchema,
    }),
    type: 'object',
    required: ['status', 'targetRevision', 'idempotencyKey', 'commandCount', 'touchedRanges', 'touchedCellCount'],
    additionalProperties: false,
    properties: {
      status: { enum: ['accepted', 'previewed', 'applied', 'rejected', 'noop'] },
      bundleId: exactString,
      targetRevision: nonNegativeSafeInteger,
      idempotencyKey: exactString,
      commandCount: nonNegativeSafeInteger,
      touchedRanges: { type: 'array', items: cellRange },
      touchedCellCount: nonNegativeSafeInteger,
      receipts: { type: 'array', items: { $ref: '#/$defs/receipt' } },
      matched: { oneOf: [{ type: 'boolean' }, { type: 'null' }] },
      changedRanges: { type: 'array', items: cellRange },
      revision: nonNegativeSafeInteger,
      undo: { $ref: '#/$defs/undo' },
      errors: { type: 'array', items: exactString },
    },
    allOf: [
      {
        if: { properties: { status: { const: 'accepted' } }, required: ['status'] },
        // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
        then: {
          not: {
            anyOf: [
              { required: ['receipts'] },
              { required: ['matched'] },
              { required: ['changedRanges'] },
              { required: ['revision'] },
              { required: ['undo'] },
              { required: ['errors'] },
            ],
          },
        },
      },
      {
        if: {
          properties: { status: { enum: ['previewed', 'applied', 'rejected', 'noop'] } },
          required: ['status'],
        },
        // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
        then: {
          required: ['receipts', 'matched', 'changedRanges'],
          properties: {
            receipts: { type: 'array', minItems: 1, items: { $ref: '#/$defs/receipt' } },
          },
        },
      },
    ],
  }),

  runResult: schema('runResult', {
    $defs: defs({
      checkExpectation: checkExpectationSchema,
      apply: applySummarySchema,
      applyCommandReceipt: applyCommandReceiptSchema,
      check: checkDataSchema,
      change: changeDataSchema,
      error: {
        type: 'object',
        required: ['code', 'message'],
        additionalProperties: false,
        properties: {
          code: { enum: workbookRunErrorCodes },
          message: { type: 'string', minLength: 1 },
          path: { type: 'string' },
          issueCode: { type: 'string' },
        },
      },
      undo: workbookUndoRefSchema,
      unverified: unverifiedDataSchema,
    }),
    oneOf: [
      {
        type: 'object',
        required: ['status', 'changed', 'checks'],
        properties: {
          status: { const: 'done' },
          apply: { $ref: '#/$defs/apply' },
          changed: { type: 'array', items: { $ref: '#/$defs/change' } },
          checks: { type: 'array', items: { $ref: '#/$defs/check' } },
          undo: { $ref: '#/$defs/undo' },
          unverified: { type: 'array', items: { $ref: '#/$defs/unverified' } },
        },
        additionalProperties: false,
      },
      {
        type: 'object',
        required: ['status', 'errors', 'changed', 'checks'],
        properties: {
          status: { const: 'failed' },
          errors: { type: 'array', items: { $ref: '#/$defs/error' } },
          apply: { $ref: '#/$defs/apply' },
          changed: { type: 'array', items: { $ref: '#/$defs/change' } },
          checks: { type: 'array', items: { $ref: '#/$defs/check' } },
          undo: { $ref: '#/$defs/undo' },
          unverified: { type: 'array', items: { $ref: '#/$defs/unverified' } },
        },
        additionalProperties: false,
      },
    ],
  }),

  readbackProof: schema('readbackProof', {
    $defs: defs({
      checkExpectation: checkExpectationSchema,
      check: checkDataSchema,
      readback: {
        type: 'object',
        required: ['target'],
        additionalProperties: false,
        properties: {
          target: refData,
          value: jsonValue,
          formula: { oneOf: [{ type: 'string' }, { type: 'null' }] },
          formulaLabels: {
            type: 'array',
            items: {
              type: 'object',
              required: ['name', 'source'],
              additionalProperties: false,
              properties: {
                name: { type: 'string', minLength: 1 },
                source: { type: 'string' },
              },
            },
          },
        },
      },
    }),
    type: 'object',
    required: ['checks', 'readbacks'],
    additionalProperties: false,
    properties: {
      checks: { type: 'array', items: { $ref: '#/$defs/check' } },
      readbacks: { type: 'array', items: { $ref: '#/$defs/readback' } },
    },
  }),
} satisfies Record<WorkbookJsonSchemaName, WorkbookJsonSchema>)

export const workbookJsonSchemaHashes = createSchemaHashes()

export const workbookJsonSchemaBundleHash = workbookJsonSchemaHash({
  version: workbookJsonSchemaVersion,
  schemas: workbookJsonSchemas,
})

export function workbookJsonSchemaHash(schemaValue: WorkbookJsonSchemaValue): string {
  const json = JSON.stringify(canonicalValue(schemaValue))
  return `bilig-schema-v1:${fnv1a64(json, 0xcbf29ce484222325n)}${fnv1a64(json, 0x84222325cbf29cen)}`
}

function canonicalValue(value: WorkbookJsonSchemaValue): WorkbookJsonSchemaValue {
  if (Array.isArray(value)) {
    return value.map(canonicalValue)
  }
  if (typeof value === 'object' && value !== null) {
    return Object.fromEntries(
      Object.entries(value)
        .toSorted(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalValue(entry)]),
    )
  }
  return value
}

function createSchemaHashes(): Readonly<Record<WorkbookJsonSchemaName, string>> {
  return deepFreeze({
    refData: workbookJsonSchemaHash(workbookJsonSchemas.refData),
    modelDescription: workbookJsonSchemaHash(workbookJsonSchemas.modelDescription),
    planData: workbookJsonSchemaHash(workbookJsonSchemas.planData),
    runtimeRequirements: workbookJsonSchemaHash(workbookJsonSchemas.runtimeRequirements),
    commandBundle: workbookJsonSchemaHash(workbookJsonSchemas.commandBundle),
    commandResult: workbookJsonSchemaHash(workbookJsonSchemas.commandResult),
    runResult: workbookJsonSchemaHash(workbookJsonSchemas.runResult),
    readbackProof: workbookJsonSchemaHash(workbookJsonSchemas.readbackProof),
  })
}

function fnv1a64(input: string, seed: bigint): string {
  let hash = seed
  for (let index = 0; index < input.length; index += 1) {
    hash ^= BigInt(input.charCodeAt(index))
    hash = BigInt.asUintN(64, hash * 0x100000001b3n)
  }
  return hash.toString(16).padStart(16, '0')
}

function deepFreeze<T>(value: T, seen = new WeakSet<object>()): T {
  if (typeof value !== 'object' || value === null || seen.has(value)) {
    return value
  }
  seen.add(value)
  for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(value))) {
    if ('value' in descriptor) {
      deepFreeze(descriptor.value, seen)
    }
  }
  return Object.freeze(value)
}
