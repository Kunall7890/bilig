import type { WorkbookJsonSchemaValue } from './schema.js'

const engineOp = Object.freeze({ $ref: '#/$defs/engineOp' })
const cellRange = Object.freeze({ $ref: '#/$defs/cellRange' })
const actionInput = Object.freeze({ $ref: '#/$defs/actionInput' })

function styleSectionRequestSchema(section: string, keys: readonly string[]): WorkbookJsonSchemaValue {
  return {
    required: [section],
    properties: {
      [section]: {
        oneOf: [
          { type: 'null' },
          {
            type: 'object',
            anyOf: keys.map((key) => ({ required: [key] })),
          },
        ],
      },
    },
  }
}

function bordersStyleRequestSchema(): WorkbookJsonSchemaValue {
  const borderSideRequestSchema: WorkbookJsonSchemaValue = {
    oneOf: [
      { type: 'null' },
      {
        type: 'object',
        anyOf: [{ required: ['style'] }, { required: ['weight'] }, { required: ['color'] }],
      },
    ],
  }

  return {
    required: ['borders'],
    properties: {
      borders: {
        oneOf: [
          { type: 'null' },
          {
            type: 'object',
            anyOf: ['top', 'right', 'bottom', 'left'].map((side) => ({
              required: [side],
              properties: {
                [side]: borderSideRequestSchema,
              },
            })),
          },
        ],
      },
    },
  }
}

export const workbookActionCellStylePatchSchema: WorkbookJsonSchemaValue = {
  allOf: [
    { $ref: '#/$defs/cellStylePatch' },
    {
      anyOf: [
        styleSectionRequestSchema('fill', ['backgroundColor']),
        styleSectionRequestSchema('font', ['family', 'size', 'bold', 'italic', 'underline', 'color']),
        styleSectionRequestSchema('alignment', [
          'horizontal',
          'vertical',
          'wrap',
          'indent',
          'shrinkToFit',
          'readingOrder',
          'textRotation',
          'justifyLastLine',
        ]),
        bordersStyleRequestSchema(),
      ],
    },
  ],
}

export const workbookRefDataDefSchemas: Record<string, WorkbookJsonSchemaValue> = {
  rangeRefData: {
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
  nameRefData: {
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
  tableRefData: {
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
  columnRefData: {
    type: 'object',
    required: ['kind', 'id', 'label', 'table', 'name'],
    additionalProperties: false,
    properties: {
      kind: { const: 'column' },
      id: { type: 'string', minLength: 1 },
      label: { type: 'string', minLength: 1 },
      table: { $ref: '#/$defs/tableRefData' },
      rows: { $ref: '#/$defs/rowsRefData' },
      name: { type: 'string', minLength: 1 },
    },
  },
  rowsRefData: {
    type: 'object',
    required: ['kind', 'id', 'label', 'where'],
    additionalProperties: false,
    properties: {
      kind: { const: 'rows' },
      id: { type: 'string', minLength: 1 },
      label: { type: 'string', minLength: 1 },
      sheetName: { type: 'string', minLength: 1 },
      table: { $ref: '#/$defs/tableRefData' },
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
}

export const workbookRefDataSchemaRefs = Object.freeze([
  { $ref: '#/$defs/rangeRefData' },
  { $ref: '#/$defs/nameRefData' },
  { $ref: '#/$defs/tableRefData' },
  { $ref: '#/$defs/columnRefData' },
  { $ref: '#/$defs/rowsRefData' },
] as const satisfies readonly WorkbookJsonSchemaValue[])

export const workbookUndoRefSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id'],
  additionalProperties: false,
  properties: {
    id: { type: 'string', minLength: 1 },
    ops: { type: 'array', items: engineOp },
  },
}

export const workbookCommandReceiptSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['status', 'featureId', 'commandId', 'category'],
  properties: {
    status: { enum: ['previewed', 'applied', 'rejected', 'noop'] },
    featureId: { type: 'string', minLength: 1 },
    commandId: { type: 'string', minLength: 1 },
    category: { enum: ['command', 'operation', 'mutation'] },
    previewOps: { type: 'array', items: engineOp },
    appliedOps: { type: 'array', items: engineOp },
    undo: workbookUndoRefSchema,
    changedRanges: { type: 'array', items: cellRange },
    proof: actionInput,
    message: { type: 'string', minLength: 1 },
    metadata: actionInput,
    errors: { type: 'array', items: { type: 'string', minLength: 1 } },
  },
  additionalProperties: false,
  allOf: [
    {
      if: { properties: { status: { const: 'previewed' } }, required: ['status'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: {
        required: ['previewOps'],
        properties: { previewOps: { type: 'array', minItems: 1, items: engineOp } },
        not: { anyOf: [{ required: ['appliedOps'] }, { required: ['undo'] }, { required: ['errors'] }] },
      },
    },
    {
      if: { properties: { status: { const: 'applied' } }, required: ['status'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: {
        anyOf: [{ required: ['appliedOps'] }, { required: ['changedRanges'] }, { required: ['undo'] }, { required: ['proof'] }],
        not: { required: ['errors'] },
      },
    },
    {
      if: { properties: { status: { const: 'rejected' } }, required: ['status'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: {
        anyOf: [{ required: ['message'] }, { required: ['errors'], properties: { errors: { type: 'array', minItems: 1 } } }],
        not: {
          anyOf: [
            { required: ['previewOps'] },
            { required: ['appliedOps'] },
            { required: ['undo'] },
            { required: ['changedRanges'] },
            { required: ['proof'] },
          ],
        },
      },
    },
    {
      if: { properties: { status: { const: 'noop' } }, required: ['status'] },
      // oxlint-disable-next-line eslint-plugin-unicorn(no-thenable) -- JSON Schema conditional schemas use the standard "then" keyword.
      then: {
        not: {
          anyOf: [
            { required: ['undo'] },
            { required: ['errors'] },
            { required: ['previewOps'], properties: { previewOps: { type: 'array', minItems: 1 } } },
            { required: ['appliedOps'], properties: { appliedOps: { type: 'array', minItems: 1 } } },
            { required: ['changedRanges'], properties: { changedRanges: { type: 'array', minItems: 1 } } },
          ],
        },
      },
    },
  ],
}
