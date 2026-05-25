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

export type WorkbookJsonSchemaName = 'refData' | 'planData' | 'commandBundle' | 'commandResult' | 'runResult' | 'readbackProof'

export const workbookJsonSchemaVersion = 'bilig-workbook-json-schema-v1'

export const workbookJsonSchemaNames = Object.freeze([
  'refData',
  'planData',
  'commandBundle',
  'commandResult',
  'runResult',
  'readbackProof',
] as const satisfies readonly WorkbookJsonSchemaName[])

const jsonValue = Object.freeze({ $ref: '#/$defs/jsonValue' })
const refData = Object.freeze({ $ref: '#/$defs/refData' })
const cellRange = Object.freeze({ $ref: '#/$defs/cellRange' })
const engineOp = Object.freeze({ $ref: '#/$defs/engineOp' })
const actionInput = Object.freeze({ $ref: '#/$defs/actionInput' })

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
    actionInput: { $ref: '#/$defs/jsonValue' },
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
    engineOp: {
      type: 'object',
      required: ['kind'],
      properties: {
        kind: { type: 'string', minLength: 1 },
      },
      additionalProperties: true,
    },
    refData: {
      oneOf: [
        {
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
        {
          type: 'object',
          required: ['kind', 'id', 'label', 'name'],
          additionalProperties: false,
          properties: {
            kind: { const: 'name' },
            id: { type: 'string', minLength: 1 },
            label: { type: 'string', minLength: 1 },
            name: { type: 'string', minLength: 1 },
          },
        },
        {
          type: 'object',
          required: ['kind', 'id', 'label'],
          additionalProperties: false,
          properties: {
            kind: { const: 'table' },
            id: { type: 'string', minLength: 1 },
            label: { type: 'string', minLength: 1 },
            name: { type: 'string', minLength: 1 },
            sheetName: { type: 'string', minLength: 1 },
            headers: { type: 'array', items: { type: 'string' } },
          },
        },
        {
          type: 'object',
          required: ['kind', 'id', 'label', 'table', 'name'],
          additionalProperties: false,
          properties: {
            kind: { const: 'column' },
            id: { type: 'string', minLength: 1 },
            label: { type: 'string', minLength: 1 },
            table: { $ref: '#/$defs/refData' },
            rows: { $ref: '#/$defs/refData' },
            name: { type: 'string', minLength: 1 },
          },
        },
        {
          type: 'object',
          required: ['kind', 'id', 'label', 'where'],
          additionalProperties: false,
          properties: {
            kind: { const: 'rows' },
            id: { type: 'string', minLength: 1 },
            label: { type: 'string', minLength: 1 },
            sheetName: { type: 'string', minLength: 1 },
            table: { $ref: '#/$defs/refData' },
            where: {
              type: 'object',
              required: ['column', 'op', 'value'],
              additionalProperties: false,
              properties: {
                column: { type: 'string', minLength: 1 },
                op: { enum: ['eq', 'neq', 'contains', 'startsWith', 'gt', 'gte', 'lt', 'lte'] },
                value: { $ref: '#/$defs/jsonValue' },
              },
            },
          },
        },
      ],
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
    expectation: { type: 'object', additionalProperties: true },
    proof: actionInput,
  },
  additionalProperties: false,
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

export const workbookJsonSchemas = deepFreeze({
  refData: schema('refData', {
    $defs: defs(),
    ...refData,
  }),

  planData: schema('planData', {
    $defs: defs({
      formulaLabel: {
        type: 'object',
        required: ['name', 'ref'],
        additionalProperties: false,
        properties: {
          name: { type: 'string', minLength: 1 },
          ref: refData,
        },
      },
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
            properties: { kind: { const: 'writeValue' }, target: refData, value: jsonValue },
          },
          {
            type: 'object',
            required: ['kind', 'target'],
            additionalProperties: false,
            properties: {
              kind: { const: 'format' },
              target: refData,
              style: { type: 'object', additionalProperties: true },
              numberFormat: { oneOf: [{ type: 'string' }, { type: 'null' }] },
            },
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

  commandBundle: schema('commandBundle', {
    $defs: defs({
      commandRequest: {
        type: 'object',
        required: ['featureId', 'commandId'],
        additionalProperties: false,
        properties: {
          featureId: { type: 'string', minLength: 1 },
          commandId: { type: 'string', minLength: 1 },
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
              id: { type: 'string', minLength: 1 },
              touchedRanges: { type: 'array', items: cellRange },
              destructive: { type: 'boolean' },
              request: { $ref: '#/$defs/commandRequest' },
            },
          },
          {
            type: 'object',
            required: ['kind', 'op'],
            additionalProperties: false,
            properties: {
              kind: { const: 'op' },
              id: { type: 'string', minLength: 1 },
              touchedRanges: { type: 'array', items: cellRange },
              destructive: { type: 'boolean' },
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
      id: { type: 'string', minLength: 1 },
      targetRevision: { type: 'integer', minimum: 0 },
      idempotencyKey: { type: 'string', minLength: 1 },
      scope: {
        type: 'object',
        additionalProperties: false,
        properties: { maxTouchedCells: { type: 'integer', minimum: 0 } },
      },
      commands: { type: 'array', minItems: 1, items: { $ref: '#/$defs/bundleCommand' } },
    },
  }),

  commandResult: schema('commandResult', {
    $defs: defs({
      receipt: {
        type: 'object',
        required: ['status', 'featureId', 'commandId', 'category'],
        properties: {
          status: { enum: ['previewed', 'applied', 'rejected', 'noop'] },
          featureId: { type: 'string', minLength: 1 },
          commandId: { type: 'string', minLength: 1 },
          category: { enum: ['command', 'operation', 'mutation'] },
          previewOps: { type: 'array', items: engineOp },
          appliedOps: { type: 'array', items: engineOp },
          undo: { type: 'object', additionalProperties: true },
          changedRanges: { type: 'array', items: cellRange },
          proof: actionInput,
          message: { type: 'string' },
          metadata: actionInput,
          errors: { type: 'array', items: { type: 'string' } },
        },
        additionalProperties: false,
      },
    }),
    type: 'object',
    required: ['status', 'targetRevision', 'idempotencyKey', 'commandCount', 'touchedRanges', 'touchedCellCount'],
    additionalProperties: false,
    properties: {
      status: { enum: ['accepted', 'previewed', 'applied', 'rejected', 'noop'] },
      bundleId: { type: 'string', minLength: 1 },
      targetRevision: { type: 'integer', minimum: 0 },
      idempotencyKey: { type: 'string', minLength: 1 },
      commandCount: { type: 'integer', minimum: 0 },
      touchedRanges: { type: 'array', items: cellRange },
      touchedCellCount: { type: 'integer', minimum: 0 },
      receipts: { type: 'array', items: { $ref: '#/$defs/receipt' } },
      matched: { oneOf: [{ type: 'boolean' }, { type: 'null' }] },
      changedRanges: { type: 'array', items: cellRange },
      revision: { type: 'integer', minimum: 0 },
      undo: { type: 'object', additionalProperties: true },
      errors: { type: 'array', items: { type: 'string' } },
    },
  }),

  runResult: schema('runResult', {
    $defs: defs({
      check: checkDataSchema,
      change: changeDataSchema,
      error: {
        type: 'object',
        required: ['code', 'message'],
        additionalProperties: false,
        properties: {
          code: { type: 'string', minLength: 1 },
          message: { type: 'string', minLength: 1 },
          path: { type: 'string' },
          issueCode: { type: 'string' },
        },
      },
    }),
    oneOf: [
      {
        type: 'object',
        required: ['status', 'changed', 'checks'],
        properties: {
          status: { const: 'done' },
          apply: { type: 'object', additionalProperties: true },
          changed: { type: 'array', items: { $ref: '#/$defs/change' } },
          checks: { type: 'array', items: { $ref: '#/$defs/check' } },
          undo: { type: 'object', additionalProperties: true },
          unverified: { type: 'array', items: { type: 'object', additionalProperties: true } },
        },
        additionalProperties: false,
      },
      {
        type: 'object',
        required: ['status', 'errors', 'changed', 'checks'],
        properties: {
          status: { const: 'failed' },
          errors: { type: 'array', items: { $ref: '#/$defs/error' } },
          apply: { type: 'object', additionalProperties: true },
          changed: { type: 'array', items: { $ref: '#/$defs/change' } },
          checks: { type: 'array', items: { $ref: '#/$defs/check' } },
          undo: { type: 'object', additionalProperties: true },
          unverified: { type: 'array', items: { type: 'object', additionalProperties: true } },
        },
        additionalProperties: false,
      },
    ],
  }),

  readbackProof: schema('readbackProof', {
    $defs: defs({
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
    planData: workbookJsonSchemaHash(workbookJsonSchemas.planData),
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
