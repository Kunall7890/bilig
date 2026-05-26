import { workbookOpSchema } from './op-schema.js'
import type { WorkbookJsonSchemaValue } from './schema.js'

type WorkbookJsonSchemaObject = {
  readonly [key: string]: WorkbookJsonSchemaValue
}

const noopEffectCommandKinds = Object.freeze(['writeValue', 'writeFormula', 'clear', 'format', 'op'] as const)

export interface NoopEffectSchemaRefs {
  readonly literalInput: WorkbookJsonSchemaValue
  readonly exactString: WorkbookJsonSchemaValue
  readonly engineOp: WorkbookJsonSchemaValue
  readonly actionCellStylePatch: WorkbookJsonSchemaValue
}

function isSchemaObject(value: WorkbookJsonSchemaValue | undefined): value is WorkbookJsonSchemaObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function schemaConstKind(value: WorkbookJsonSchemaValue): string | undefined {
  if (!isSchemaObject(value)) {
    return undefined
  }
  const properties = value['properties']
  if (!isSchemaObject(properties)) {
    return undefined
  }
  const kind = properties['kind']
  if (!isSchemaObject(kind)) {
    return undefined
  }
  const kindConst = kind['const']
  return typeof kindConst === 'string' ? kindConst : undefined
}

function opNoopEffectSchemas(refs: NoopEffectSchemaRefs): readonly WorkbookJsonSchemaValue[] {
  if (!isSchemaObject(workbookOpSchema) || !Array.isArray(workbookOpSchema['oneOf'])) {
    return [
      {
        type: 'object',
        required: ['kind', 'opKind', 'op'],
        additionalProperties: true,
        properties: {
          kind: { const: 'op' },
          opKind: refs.exactString,
          op: refs.engineOp,
        },
      },
    ]
  }

  return workbookOpSchema['oneOf'].map((opSchema: WorkbookJsonSchemaValue) => {
    const opKind = schemaConstKind(opSchema)
    return {
      type: 'object',
      required: ['kind', 'opKind', 'op'],
      additionalProperties: true,
      properties: {
        kind: { const: 'op' },
        opKind: opKind === undefined ? refs.exactString : { const: opKind },
        op: opSchema,
      },
    }
  })
}

function noopEffectCommandKindCondition(commandKind: (typeof noopEffectCommandKinds)[number]): WorkbookJsonSchemaValue {
  return {
    if: {
      required: ['commandKind', 'noop'],
      properties: {
        commandKind: { const: commandKind },
      },
    },
    // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
    then: {
      properties: {
        noop: {
          properties: {
            proof: {
              properties: {
                commandKind: { const: commandKind },
                effect: {
                  required: ['kind'],
                  properties: {
                    kind: { const: commandKind },
                  },
                },
              },
            },
          },
        },
      },
    },
  }
}

export function createNoopEffectSchema(refs: NoopEffectSchemaRefs): WorkbookJsonSchemaValue {
  return {
    oneOf: [
      {
        type: 'object',
        required: ['kind', 'value'],
        additionalProperties: true,
        properties: {
          kind: { const: 'writeValue' },
          value: refs.literalInput,
        },
      },
      {
        type: 'object',
        required: ['kind', 'formula'],
        additionalProperties: true,
        properties: {
          kind: { const: 'writeFormula' },
          formula: { type: 'string', minLength: 1 },
        },
      },
      {
        type: 'object',
        required: ['kind', 'cleared'],
        additionalProperties: true,
        properties: {
          kind: { const: 'clear' },
          cleared: { const: true },
        },
      },
      {
        type: 'object',
        required: ['kind'],
        additionalProperties: true,
        properties: {
          kind: { const: 'format' },
          style: refs.actionCellStylePatch,
          numberFormat: { oneOf: [{ type: 'string' }, { type: 'null' }] },
        },
        anyOf: [{ required: ['style'] }, { required: ['numberFormat'] }],
      },
      ...opNoopEffectSchemas(refs),
    ],
  }
}

export function noopEffectCommandKindConditions(): readonly WorkbookJsonSchemaValue[] {
  return noopEffectCommandKinds.map(noopEffectCommandKindCondition)
}
