import {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
} from '@bilig/protocol'
import type { WorkbookJsonSchemaValue } from './schema.js'

const nonEmptyStringSchema = { type: 'string', minLength: 1 } as const
const nonNegativeIntegerSchema = { type: 'integer', minimum: 0, maximum: Number.MAX_SAFE_INTEGER } as const
const positiveIntegerSchema = { type: 'integer', minimum: 1, maximum: Number.MAX_SAFE_INTEGER } as const
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
      additionalProperties: false,
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
      additionalProperties: false,
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
      additionalProperties: false,
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
      additionalProperties: false,
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
      additionalProperties: false,
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
  additionalProperties: false,
  properties: {
    fill: styleFillPatchSchema,
    font: styleFontPatchSchema,
    alignment: styleAlignmentPatchSchema,
    borders: styleBordersPatchSchema,
  },
}

const styleFillRecordSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['backgroundColor'],
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

const workbookCalculationSettingsSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['mode'],
  additionalProperties: true,
  properties: {
    mode: { enum: ['automatic', 'manual'] },
    compatibilityMode: { enum: ['excel-modern', 'odf-1.4'] },
    dateSystem: { enum: ['1900', '1904'] },
    calcId: nullableNumberSchema,
    iterate: nullableBooleanSchema,
    iterateCount: nullableNumberSchema,
    iterateDelta: nullableStringSchema,
    fullPrecision: nullableBooleanSchema,
    fullCalcOnLoad: nullableBooleanSchema,
    forceFullCalc: nullableBooleanSchema,
    calcOnSave: nullableBooleanSchema,
    calcCompleted: nullableBooleanSchema,
    concurrentCalc: nullableBooleanSchema,
  },
}

const workbookVolatileContextSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['recalcEpoch'],
  additionalProperties: true,
  properties: {
    recalcEpoch: numberSchema,
  },
}

const sheetProtectionSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['sheetName'],
  additionalProperties: true,
  properties: {
    sheetName: nonEmptyStringSchema,
    hideFormulas: booleanSchema,
  },
}

const rangeProtectionSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'range'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    range: cellRangeRefSchema,
    hideFormulas: booleanSchema,
  },
}

const validationListSourceSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    {
      type: 'object',
      required: ['kind', 'name'],
      additionalProperties: true,
      properties: { kind: { const: 'named-range' }, name: nonEmptyStringSchema },
    },
    {
      type: 'object',
      required: ['kind', 'sheetName', 'address'],
      additionalProperties: true,
      properties: { kind: { const: 'cell-ref' }, sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema },
    },
    {
      type: 'object',
      required: ['kind', 'sheetName', 'startAddress', 'endAddress'],
      additionalProperties: true,
      properties: {
        kind: { const: 'range-ref' },
        sheetName: nonEmptyStringSchema,
        startAddress: nonEmptyStringSchema,
        endAddress: nonEmptyStringSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'tableName', 'columnName'],
      additionalProperties: true,
      properties: { kind: { const: 'structured-ref' }, tableName: nonEmptyStringSchema, columnName: nonEmptyStringSchema },
    },
    {
      type: 'object',
      required: ['kind', 'formula'],
      additionalProperties: true,
      properties: { kind: { const: 'formula' }, formula: { type: 'string' } },
    },
  ],
}

const rangeComparisonOperatorSchema = {
  enum: ['between', 'notBetween'],
} as const
const singleComparisonOperatorSchema = {
  enum: ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'],
} as const
const scalarValidationKindSchema = {
  enum: ['whole', 'decimal', 'date', 'time', 'textLength'],
} as const
const twoValueComparisonValuesSchema = {
  type: 'array',
  minItems: 2,
  maxItems: 2,
  items: literalInputRefSchema,
} as const
const oneValueComparisonValuesSchema = {
  type: 'array',
  minItems: 1,
  maxItems: 1,
  items: literalInputRefSchema,
} as const

const validationRuleSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    {
      type: 'object',
      required: ['kind'],
      additionalProperties: true,
      properties: {
        kind: { const: 'list' },
        values: { type: 'array', items: literalInputRefSchema },
        source: validationListSourceSchema,
      },
      oneOf: [
        { required: ['values'], not: { required: ['source'] } },
        { required: ['source'], not: { required: ['values'] } },
      ],
    },
    {
      type: 'object',
      required: ['kind'],
      additionalProperties: true,
      properties: {
        kind: { const: 'checkbox' },
        checkedValue: literalInputRefSchema,
        uncheckedValue: literalInputRefSchema,
      },
    },
    { type: 'object', required: ['kind'], additionalProperties: true, properties: { kind: { const: 'any' } } },
    {
      type: 'object',
      required: ['kind', 'operator', 'values'],
      additionalProperties: true,
      properties: {
        kind: scalarValidationKindSchema,
        operator: rangeComparisonOperatorSchema,
        values: twoValueComparisonValuesSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'operator', 'values'],
      additionalProperties: true,
      properties: {
        kind: scalarValidationKindSchema,
        operator: singleComparisonOperatorSchema,
        values: oneValueComparisonValuesSchema,
      },
    },
  ],
}

const dataValidationSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['range', 'rule'],
  additionalProperties: true,
  properties: {
    range: cellRangeRefSchema,
    rule: validationRuleSchema,
    allowBlank: booleanSchema,
    showDropdown: booleanSchema,
    promptTitle: { type: 'string' },
    promptMessage: { type: 'string' },
    errorStyle: { enum: ['stop', 'warning', 'information'] },
    errorTitle: { type: 'string' },
    errorMessage: { type: 'string' },
  },
}

const conditionalFormatRuleSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    {
      type: 'object',
      required: ['kind', 'operator', 'values'],
      additionalProperties: true,
      properties: {
        kind: { const: 'cellIs' },
        operator: rangeComparisonOperatorSchema,
        values: twoValueComparisonValuesSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'operator', 'values'],
      additionalProperties: true,
      properties: {
        kind: { const: 'cellIs' },
        operator: singleComparisonOperatorSchema,
        values: oneValueComparisonValuesSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'text'],
      additionalProperties: true,
      properties: {
        kind: { const: 'textContains' },
        text: { type: 'string' },
        caseSensitive: booleanSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'formula'],
      additionalProperties: true,
      properties: { kind: { const: 'formula' }, formula: { type: 'string' } },
    },
    { type: 'object', required: ['kind'], additionalProperties: true, properties: { kind: { const: 'blanks' } } },
    { type: 'object', required: ['kind'], additionalProperties: true, properties: { kind: { const: 'notBlanks' } } },
  ],
}

const conditionalFormatSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'range', 'rule', 'style'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    range: cellRangeRefSchema,
    rule: conditionalFormatRuleSchema,
    style: cellStylePatchSchema,
    stopIfTrue: booleanSchema,
    priority: nonNegativeIntegerSchema,
  },
}

const conditionalFormatArtifactsSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['xml'],
  additionalProperties: true,
  properties: {
    xml: { type: 'string' },
  },
}

const commentEntrySchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'body'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    body: { type: 'string' },
    authorUserId: { type: 'string' },
    authorDisplayName: { type: 'string' },
    createdAtUnixMs: nonNegativeIntegerSchema,
  },
}

const commentThreadSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['threadId', 'sheetName', 'address', 'comments'],
  additionalProperties: true,
  properties: {
    threadId: nonEmptyStringSchema,
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    comments: { type: 'array', minItems: 1, items: commentEntrySchema },
    resolved: booleanSchema,
    resolvedByUserId: { type: 'string' },
    resolvedAtUnixMs: nonNegativeIntegerSchema,
  },
}

const noteSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['sheetName', 'address', 'text'],
  additionalProperties: true,
  properties: {
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    text: { type: 'string' },
  },
}

const hyperlinkSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['sheetName', 'address', 'target'],
  additionalProperties: true,
  properties: {
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    target: { type: 'string' },
    tooltip: { type: 'string' },
    display: { type: 'string' },
  },
}

const definedNameValueSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    literalInputRefSchema,
    {
      type: 'object',
      required: ['kind', 'value'],
      additionalProperties: true,
      properties: { kind: { const: 'scalar' }, value: literalInputRefSchema },
    },
    {
      type: 'object',
      required: ['kind', 'sheetName', 'address'],
      additionalProperties: true,
      properties: { kind: { const: 'cell-ref' }, sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema },
    },
    {
      type: 'object',
      required: ['kind', 'sheetName', 'startAddress', 'endAddress'],
      additionalProperties: true,
      properties: {
        kind: { const: 'range-ref' },
        sheetName: nonEmptyStringSchema,
        startAddress: nonEmptyStringSchema,
        endAddress: nonEmptyStringSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'tableName', 'columnName'],
      additionalProperties: true,
      properties: { kind: { const: 'structured-ref' }, tableName: nonEmptyStringSchema, columnName: nonEmptyStringSchema },
    },
    {
      type: 'object',
      required: ['kind', 'formula'],
      additionalProperties: true,
      properties: { kind: { const: 'formula' }, formula: { type: 'string' } },
    },
  ],
}

const autoFilterValueCriteriaSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['values'],
  additionalProperties: true,
  properties: {
    blank: booleanSchema,
    values: { type: 'array', items: { type: 'string' } },
  },
}

const autoFilterCustomCriterionSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['value'],
  additionalProperties: true,
  properties: {
    operator: { enum: ['equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan'] },
    value: { type: 'string' },
  },
}

const autoFilterCustomCriteriaSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['filters'],
  additionalProperties: true,
  properties: {
    and: booleanSchema,
    filters: { type: 'array', items: autoFilterCustomCriterionSchema },
  },
}

const autoFilterColumnSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['colId'],
  additionalProperties: true,
  properties: {
    colId: nonNegativeIntegerSchema,
    hiddenButton: booleanSchema,
    showButton: booleanSchema,
    filters: autoFilterValueCriteriaSchema,
    customFilters: autoFilterCustomCriteriaSchema,
  },
}

const autoFilterSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['sheetName', 'startAddress', 'endAddress'],
  additionalProperties: true,
  properties: {
    sheetName: nonEmptyStringSchema,
    startAddress: nonEmptyStringSchema,
    endAddress: nonEmptyStringSchema,
    criteria: { type: 'array', items: autoFilterColumnSchema },
  },
}

const tableColumnSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['name'],
  additionalProperties: true,
  properties: {
    name: nonEmptyStringSchema,
    calculatedColumnFormula: { type: 'string' },
    totalsRowLabel: { type: 'string' },
    totalsRowFunction: { type: 'string' },
    totalsRowFormula: { type: 'string' },
  },
}

const tableStyleSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  additionalProperties: true,
  properties: {
    name: { type: 'string' },
    showFirstColumn: booleanSchema,
    showLastColumn: booleanSchema,
    showRowStripes: booleanSchema,
    showColumnStripes: booleanSchema,
  },
}

const tableSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['name', 'sheetName', 'startAddress', 'endAddress', 'columnNames', 'headerRow', 'totalsRow'],
  additionalProperties: true,
  properties: {
    name: nonEmptyStringSchema,
    sheetName: nonEmptyStringSchema,
    startAddress: nonEmptyStringSchema,
    endAddress: nonEmptyStringSchema,
    columnNames: { type: 'array', items: { type: 'string' } },
    columns: { type: 'array', items: tableColumnSchema },
    headerRow: booleanSchema,
    totalsRow: booleanSchema,
    style: tableStyleSchema,
    autoFilter: autoFilterSchema,
    sortState: { type: 'string' },
  },
}

const pivotValueSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['sourceColumn', 'summarizeBy'],
  additionalProperties: true,
  properties: {
    sourceColumn: nonEmptyStringSchema,
    summarizeBy: { enum: ['sum', 'count'] },
    outputLabel: { type: 'string' },
  },
}

const drawingAnchorMarkerSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['row', 'col'],
  additionalProperties: true,
  properties: {
    row: nonNegativeIntegerSchema,
    col: nonNegativeIntegerSchema,
    rowOffset: nonNegativeIntegerSchema,
    colOffset: nonNegativeIntegerSchema,
  },
}

const drawingAnchorExtentSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['width', 'height'],
  additionalProperties: true,
  properties: {
    width: positiveIntegerSchema,
    height: positiveIntegerSchema,
  },
}

const drawingAnchorPositionSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['x', 'y'],
  additionalProperties: true,
  properties: {
    x: nonNegativeIntegerSchema,
    y: nonNegativeIntegerSchema,
  },
}

const chartAnchorSchema: WorkbookJsonSchemaValue = {
  oneOf: [
    {
      type: 'object',
      required: ['kind', 'from', 'to'],
      additionalProperties: true,
      properties: {
        kind: { const: 'twoCell' },
        editAs: { enum: ['twoCell', 'oneCell', 'absolute'] },
        from: drawingAnchorMarkerSchema,
        to: drawingAnchorMarkerSchema,
      },
    },
    {
      type: 'object',
      required: ['kind', 'from', 'extent'],
      additionalProperties: true,
      properties: { kind: { const: 'oneCell' }, from: drawingAnchorMarkerSchema, extent: drawingAnchorExtentSchema },
    },
    {
      type: 'object',
      required: ['kind', 'position', 'extent'],
      additionalProperties: true,
      properties: { kind: { const: 'absolute' }, position: drawingAnchorPositionSchema, extent: drawingAnchorExtentSchema },
    },
  ],
}

const chartSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'sheetName', 'address', 'source', 'chartType', 'rows', 'cols'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    source: cellRangeRefSchema,
    chartType: { enum: ['column', 'bar', 'line', 'area', 'pie', 'scatter'] },
    anchor: chartAnchorSchema,
    seriesOrientation: { enum: ['rows', 'columns'] },
    firstRowAsHeaders: booleanSchema,
    firstColumnAsLabels: booleanSchema,
    title: { type: 'string' },
    legendPosition: { enum: ['top', 'right', 'bottom', 'left', 'hidden'] },
    rows: positiveIntegerSchema,
    cols: positiveIntegerSchema,
  },
}

const imageSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'sheetName', 'address', 'sourceUrl', 'rows', 'cols'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    sourceUrl: { type: 'string' },
    rows: positiveIntegerSchema,
    cols: positiveIntegerSchema,
    altText: { type: 'string' },
  },
}

const shapeSchema: WorkbookJsonSchemaValue = {
  type: 'object',
  required: ['id', 'sheetName', 'address', 'shapeType', 'rows', 'cols'],
  additionalProperties: true,
  properties: {
    id: nonEmptyStringSchema,
    sheetName: nonEmptyStringSchema,
    address: nonEmptyStringSchema,
    shapeType: { enum: ['rectangle', 'roundedRectangle', 'ellipse', 'line', 'arrow', 'textBox'] },
    rows: positiveIntegerSchema,
    cols: positiveIntegerSchema,
    text: { type: 'string' },
    fillColor: { type: 'string' },
    strokeColor: { type: 'string' },
  },
}

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
    opSchema('setCalculationSettings', ['settings'], { settings: workbookCalculationSettingsSchema }),
    opSchema('setVolatileContext', ['context'], { context: workbookVolatileContextSchema }),
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
    opSchema('setSheetProtection', ['protection'], { protection: sheetProtectionSchema }),
    opSchema('clearSheetProtection', ['sheetName'], { sheetName: nonEmptyStringSchema }),
    opSchema('setFilter', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('clearFilter', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('setSort', ['sheetName', 'range', 'keys'], {
      sheetName: nonEmptyStringSchema,
      range: cellRangeRefSchema,
      keys: { type: 'array', items: sortKeySchema },
    }),
    opSchema('clearSort', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('setDataValidation', ['validation'], { validation: dataValidationSchema }),
    opSchema('clearDataValidation', ['sheetName', 'range'], { sheetName: nonEmptyStringSchema, range: cellRangeRefSchema }),
    opSchema('upsertConditionalFormat', ['format'], { format: conditionalFormatSchema }),
    opSchema('deleteConditionalFormat', ['id', 'sheetName'], { id: nonEmptyStringSchema, sheetName: nonEmptyStringSchema }),
    opSchema('setConditionalFormatArtifacts', ['sheetName', 'artifacts'], {
      sheetName: nonEmptyStringSchema,
      artifacts: conditionalFormatArtifactsSchema,
    }),
    opSchema('clearConditionalFormatArtifacts', ['sheetName'], { sheetName: nonEmptyStringSchema }),
    opSchema('upsertRangeProtection', ['protection'], { protection: rangeProtectionSchema }),
    opSchema('deleteRangeProtection', ['id', 'sheetName'], { id: nonEmptyStringSchema, sheetName: nonEmptyStringSchema }),
    opSchema('upsertCommentThread', ['thread'], { thread: commentThreadSchema }),
    opSchema('deleteCommentThread', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertNote', ['note'], { note: noteSchema }),
    opSchema('deleteNote', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertHyperlink', ['hyperlink'], { hyperlink: hyperlinkSchema }),
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
      value: definedNameValueSchema,
    }),
    opSchema('deleteDefinedName', ['name'], { name: nonEmptyStringSchema }),
    opSchema('upsertTable', ['table'], { table: tableSchema }),
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
      values: { type: 'array', items: pivotValueSchema },
      rows: positiveIntegerSchema,
      cols: positiveIntegerSchema,
    }),
    opSchema('deletePivotTable', ['sheetName', 'address'], { sheetName: nonEmptyStringSchema, address: nonEmptyStringSchema }),
    opSchema('upsertChart', ['chart'], { chart: chartSchema }),
    opSchema('deleteChart', ['id'], { id: nonEmptyStringSchema }),
    opSchema('upsertImage', ['image'], { image: imageSchema }),
    opSchema('deleteImage', ['id'], { id: nonEmptyStringSchema }),
    opSchema('upsertShape', ['shape'], { shape: shapeSchema }),
    opSchema('deleteShape', ['id'], { id: nonEmptyStringSchema }),
  ],
}
