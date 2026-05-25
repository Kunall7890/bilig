import {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
} from '@bilig/protocol'
import type { WorkbookJsonSchemaValue } from './schema.js'

const nonEmptyStringSchema = { type: 'string', minLength: 1 } as const
const nonNegativeIntegerSchema = { type: 'integer', minimum: 0 } as const
const positiveIntegerSchema = { type: 'integer', minimum: 1 } as const
const booleanSchema = { type: 'boolean' } as const
const numberSchema = { type: 'number' } as const
const cellRangeRefSchema = { $ref: '#/$defs/cellRange' } as const
const literalInputRefSchema = { $ref: '#/$defs/literalInput' } as const

export const literalInputSchema: WorkbookJsonSchemaValue = {
  oneOf: [{ type: 'null' }, booleanSchema, numberSchema, { type: 'string' }],
}

const nullableStringSchema = { oneOf: [{ type: 'string' }, { type: 'null' }] } as const
const nullableNumberSchema = { oneOf: [numberSchema, { type: 'null' }] } as const
const nullableBooleanSchema = { oneOf: [booleanSchema, { type: 'null' }] } as const

const styleFillPatchSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    { type: 'null' },
    {
      type: 'object',
      additionalProperties: true,
      properties: {
        backgroundColor: nullableStringSchema,
      },
    },
  ],
}

const styleFontPatchSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    { type: 'null' },
    {
      type: 'object',
      additionalProperties: true,
      properties: {
        family: nullableStringSchema,
        size: nullableNumberSchema,
        bold: nullableBooleanSchema,
        italic: nullableBooleanSchema,
        underline: nullableBooleanSchema,
        color: nullableStringSchema,
      },
    },
  ],
}

const styleAlignmentPatchSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    { type: 'null' },
    {
      type: 'object',
      additionalProperties: true,
      properties: {
        horizontal: { oneOf: [{ enum: CELL_HORIZONTAL_ALIGNMENT_VALUES }, { type: 'null' }] },
        vertical: { oneOf: [{ enum: CELL_VERTICAL_ALIGNMENT_VALUES }, { type: 'null' }] },
        wrap: nullableBooleanSchema,
        indent: nullableNumberSchema,
        shrinkToFit: nullableBooleanSchema,
        readingOrder: nullableNumberSchema,
        textRotation: nullableNumberSchema,
        justifyLastLine: nullableBooleanSchema,
      },
    },
  ],
}

const styleBorderSidePatchSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    { type: 'null' },
    {
      type: 'object',
      additionalProperties: true,
      properties: {
        style: { oneOf: [{ enum: CELL_BORDER_STYLE_VALUES }, { type: 'null' }] },
        weight: { oneOf: [{ enum: CELL_BORDER_WEIGHT_VALUES }, { type: 'null' }] },
        color: nullableStringSchema,
      },
    },
  ],
}

const styleBordersPatchSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    { type: 'null' },
    {
      type: 'object',
      additionalProperties: true,
      properties: {
        top: styleBorderSidePatchSchema,
        right: styleBorderSidePatchSchema,
        bottom: styleBorderSidePatchSchema,
        left: styleBorderSidePatchSchema,
      },
    },
  ],
}

export const cellStylePatchSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    fill: styleFillPatchSchema,
    font: styleFontPatchSchema,
    alignment: styleAlignmentPatchSchema,
    borders: styleBordersPatchSchema,
  },
}

const styleFillRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    backgroundColor: { type: 'string' },
  },
}

const styleFontRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    family: { type: 'string' },
    size: numberSchema,
    bold: booleanSchema,
    italic: booleanSchema,
    underline: booleanSchema,
    color: { type: 'string' },
  },
}

const styleAlignmentRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    horizontal: { enum: CELL_HORIZONTAL_ALIGNMENT_VALUES },
    vertical: { enum: CELL_VERTICAL_ALIGNMENT_VALUES },
    wrap: booleanSchema,
    indent: numberSchema,
    shrinkToFit: booleanSchema,
    readingOrder: numberSchema,
    textRotation: numberSchema,
    justifyLastLine: booleanSchema,
  },
}

const styleBorderSideRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['style', 'weight', 'color'],
  additionalProperties: true,
  properties: {
    style: { enum: CELL_BORDER_STYLE_VALUES },
    weight: { enum: CELL_BORDER_WEIGHT_VALUES },
    color: { type: 'string' },
  },
}

const styleBordersRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    top: styleBorderSideRecordSchema,
    right: styleBorderSideRecordSchema,
    bottom: styleBorderSideRecordSchema,
    left: styleBorderSideRecordSchema,
  },
}

const cellStyleRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    fill: styleFillRecordSchema,
    font: styleFontRecordSchema,
    alignment: styleAlignmentRecordSchema,
    borders: styleBordersRecordSchema,
    protection: {
      type: 'object',
      additionalProperties: true,
      properties: {
        locked: booleanSchema,
        hidden: booleanSchema,
      },
    },
  },
}

const cellNumberFormatRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'code', 'kind'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    code: { type: 'string' },
    kind: { enum: ['general', 'number', 'currency', 'accounting', 'percent', 'date', 'time', 'datetime', 'text'] },
  },
}

const axisEntrySchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'index'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    index: nonNegativeIntegerSchema,
    size: { oneOf: [positiveIntegerSchema, { type: 'null' }] },
    hidden: nullableBooleanSchema,
    filterHidden: nullableBooleanSchema,
  },
}

const sortKeySchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['keyAddress', 'direction'],
  additionalProperties: true,
  properties: {
    keyAddress: nonEmptyStringSchema,
    direction: { enum: ['asc', 'desc'] },
  },
}

const looseObjectSchema = { type: 'object', additionalProperties: true } as const

function opSchema(kind: string, required: readonly string[], properties: Record<string, WorkbookJsonSchemaValue>): WorkbookJsonSchemaValue {
  return {
    type: 'object',
    required: ['kind', ...required],
    additionalProperties: true,
    properties: {
      kind: { const: kind },
      ...properties,
    },
  }
}

export const workbookOpSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    opSchema('upsertWorkbook', ['name'], { name: nonEmptyStringSchema }),
    opSchema('setWorkbookMetadata', ['key', 'value'], { key: nonEmptyStringSchema, value: literalInputRefSchema }),
    opSchema('setCalculationSettings', ['settings'], { settings: looseObjectSchema }),
    opSchema('setVolatileContext', ['context'], { context: looseObjectSchema }),
    opSchema('upsertSheet', ['name', 'order'], { name: nonEmptyStringSchema, order: nonNegativeIntegerSchema, id: positiveIntegerSchema }),
    opSchema('renameSheet', ['oldName', 'newName'], { oldName: nonEmptyStringSchema, newName: nonEmptyStringSchema }),
    opSchema('deleteSheet', ['name'], { name: nonEmptyStringSchema }),
    opSchema('insertRows', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      entries: { type: 'array', items: axisEntrySchema },
    }),
    opSchema('deleteRows', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
    }),
    opSchema('moveRows', ['sheetName', 'start', 'count', 'target'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      target: nonNegativeIntegerSchema,
    }),
    opSchema('insertColumns', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      entries: { type: 'array', items: axisEntrySchema },
    }),
    opSchema('deleteColumns', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
    }),
    opSchema('moveColumns', ['sheetName', 'start', 'count', 'target'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      target: nonNegativeIntegerSchema,
    }),
    opSchema('updateRowMetadata', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      size: { oneOf: [positiveIntegerSchema, { type: 'null' }] },
      hidden: nullableBooleanSchema,
      filterHidden: nullableBooleanSchema,
    }),
    opSchema('updateColumnMetadata', ['sheetName', 'start', 'count'], {
      sheetName: nonEmptyStringSchema,
      start: nonNegativeIntegerSchema,
      count: positiveIntegerSchema,
      size: { oneOf: [positiveIntegerSchema, { type: 'null' }] },
      hidden: nullableBooleanSchema,
    }),
    opSchema('setFreezePane', ['sheetName', 'rows', 'cols'], {
      sheetName: nonEmptyStringSchema,
      rows: nonNegativeIntegerSchema,
      cols: nonNegativeIntegerSchema,
    }),
    opSchema('clearFreezePane', ['sheetName'], { sheetName: nonEmptyStringSchema }),
    opSchema('mergeCells', ['range'], { range: cellRangeRefSchema }),
    opSchema('unmergeCells', ['range'], { range: cellRangeRefSchema }),
    opSchema('setSheetProtection', ['protection'], { protection: looseObjectSchema }),
    opSchema('clearSheetProtection', ['sheetName'], { sheetName: nonEmptyStringSchema }),
    opSchema('setFilter', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('clearFilter', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('setSort', ['sheetName', 'range', 'keys'], {
      sheetName: nonEmptyStringSchema,
      range: cellRangeRefSchema,
      keys: { type: 'array', items: sortKeySchema },
    }),
    opSchema('clearSort', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('setDataValidation', ['validation'], { validation: looseObjectSchema }),
    opSchema('clearDataValidation', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('upsertConditionalFormat', ['format'], { format: looseObjectSchema }),
    opSchema('deleteConditionalFormat', ['id', 'sheetName'], { id: nonEmptyStringSchema, sheetName: nonEmptyStringSchema }),
    opSchema('setConditionalFormatArtifacts', ['sheetName', 'artifacts'], {
      sheetName: nonEmptyStringSchema,
      artifacts: looseObjectSchema,
    }),
    opSchema('clearConditionalFormatArtifacts', ['sheetName'], { sheetName: nonEmptyStringSchema }),
    opSchema('upsertRangeProtection', ['protection'], { protection: looseObjectSchema }),
    opSchema('deleteRangeProtection', ['id', 'sheetName'], { id: nonEmptyStringSchema, sheetName: nonEmptyStringSchema }),
    opSchema('upsertCommentThread', ['thread'], { thread: looseObjectSchema }),
    opSchema('deleteCommentThread', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertNote', ['note'], { note: looseObjectSchema }),
    opSchema('deleteNote', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertHyperlink', ['hyperlink'], { hyperlink: looseObjectSchema }),
    opSchema('deleteHyperlink', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('setCellValue', ['sheetName', 'address', 'value'], {
      sheetName: nonEmptyStringSchema,
      address: nonEmptyStringSchema,
      value: literalInputRefSchema,
      authoredBlank: booleanSchema,
    }),
    opSchema('setCellFormula', ['sheetName', 'address', 'formula'], {
      sheetName: nonEmptyStringSchema,
      address: nonEmptyStringSchema,
      formula: { type: 'string' },
    }),
    opSchema('setCellFormat', ['sheetName', 'address', 'format'], {
      sheetName: nonEmptyStringSchema,
      address: nonEmptyStringSchema,
      format: { oneOf: [{ type: 'string' }, { type: 'null' }] },
    }),
    opSchema('upsertCellStyle', ['style'], { style: cellStyleRecordSchema }),
    opSchema('upsertCellNumberFormat', ['format'], { format: cellNumberFormatRecordSchema }),
    opSchema('setStyleRange', ['range', 'styleId'], { range: cellRangeRefSchema, styleId: nonEmptyStringSchema }),
    opSchema('setFormatRange', ['range', 'formatId'], { range: cellRangeRefSchema, formatId: nonEmptyStringSchema }),
    opSchema('clearCell', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertDefinedName', ['name', 'value'], {
      name: nonEmptyStringSchema,
      value: { oneOf: [literalInputRefSchema, looseObjectSchema] },
    }),
    opSchema('deleteDefinedName', ['name'], { name: nonEmptyStringSchema }),
    opSchema('upsertTable', ['table'], { table: looseObjectSchema }),
    opSchema('deleteTable', ['name'], { name: nonEmptyStringSchema }),
    opSchema('upsertSpillRange', ['sheetName', 'address', 'rows', 'cols'], {
      sheetName: nonEmptyStringSchema,
      address: nonEmptyStringSchema,
      rows: positiveIntegerSchema,
      cols: positiveIntegerSchema,
    }),
    opSchema('deleteSpillRange', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertPivotTable', ['name', 'sheetName', 'address', 'source', 'groupBy', 'values', 'rows', 'cols'], {
      name: nonEmptyStringSchema,
      sheetName: nonEmptyStringSchema,
      address: nonEmptyStringSchema,
      source: cellRangeRefSchema,
      groupBy: { type: 'array', items: { type: 'string' } },
      values: { type: 'array', items: looseObjectSchema },
      rows: positiveIntegerSchema,
      cols: positiveIntegerSchema,
    }),
    opSchema('deletePivotTable', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertChart', ['chart'], { chart: looseObjectSchema }),
    opSchema('deleteChart', ['id'], { id: nonEmptyStringSchema }),
    opSchema('upsertImage', ['image'], { image: looseObjectSchema }),
    opSchema('deleteImage', ['id'], { id: nonEmptyStringSchema }),
    opSchema('upsertShape', ['shape'], { shape: looseObjectSchema }),
    opSchema('deleteShape', ['id'], { id: nonEmptyStringSchema }),
  ],
}
