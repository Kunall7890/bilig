import { normalizeFormulaFunctionName, parseFormula } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookRef } from './find.js'

export interface WorkbookFormulaExpression {
  readonly kind: 'formula'
  readonly source: string
  readonly inputs: readonly WorkbookRef[]
  readonly labels: readonly WorkbookFormulaLabel[]
}

export interface WorkbookFormulaLabel {
  readonly name: string
  readonly ref: WorkbookRef
}

export interface WorkbookRawFormulaOptions {
  readonly inputs?: readonly WorkbookRef[]
  readonly labels?: readonly WorkbookFormulaLabel[]
}

export type WorkbookFormulaOperand = WorkbookFormulaExpression | WorkbookRef | number | boolean

function normalizeFormulaSource(source: string): string {
  const trimmed = source.trim()
  if (trimmed === '') {
    throw new Error('Formula source cannot be empty')
  }
  const normalized = trimmed.startsWith('=') ? trimmed.slice(1) : trimmed
  parseFormula(normalized)
  return normalized
}

function uniqueRefs(refs: readonly WorkbookRef[]): readonly WorkbookRef[] {
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  for (const ref of refs) {
    const key = `${ref.kind}:${ref.id}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(ref)
  }
  return unique
}

function uniqueLabels(labels: readonly WorkbookFormulaLabel[]): readonly WorkbookFormulaLabel[] {
  const seen = new Set<string>()
  const unique: WorkbookFormulaLabel[] = []
  for (const label of labels) {
    if (label.name.trim() === '') {
      throw new Error('Formula label name cannot be empty')
    }
    const key = `${label.name}:${label.ref.kind}:${label.ref.id}`
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    unique.push(label)
  }
  return unique
}

function labelForRef(ref: WorkbookRef): WorkbookFormulaLabel {
  return {
    name: refSource(ref),
    ref,
  }
}

function labelsForRefs(refs: readonly WorkbookRef[]): readonly WorkbookFormulaLabel[] {
  return refs.map(labelForRef)
}

function createFormulaExpression(
  source: string,
  inputs: readonly WorkbookRef[] = [],
  labels: readonly WorkbookFormulaLabel[] = labelsForRefs(inputs),
): WorkbookFormulaExpression {
  const expressionLabels = uniqueLabels(labels)
  return {
    kind: 'formula',
    source: normalizeFormulaSource(source),
    inputs: uniqueRefs([...inputs, ...expressionLabels.map((label) => label.ref)]),
    labels: expressionLabels,
  }
}

function isFormulaExpression(value: unknown): value is WorkbookFormulaExpression {
  return typeof value === 'object' && value !== null && 'kind' in value && value.kind === 'formula'
}

function isWorkbookRef(value: unknown): value is WorkbookRef {
  return (
    typeof value === 'object' &&
    value !== null &&
    'kind' in value &&
    typeof value.kind === 'string' &&
    value.kind !== 'formula' &&
    'id' in value &&
    typeof value.id === 'string'
  )
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replaceAll("'", "''")}'`
}

function formatRangeRef(range: CellRangeRef): string {
  const sheet = quoteSheetName(range.sheetName)
  return range.startAddress === range.endAddress ? `${sheet}!${range.startAddress}` : `${sheet}!${range.startAddress}:${range.endAddress}`
}

function placeholderRefSource(ref: WorkbookRef): string {
  return `__bilig_ref_${ref.id.replaceAll(/[^A-Za-z0-9_]+/g, '_')}`
}

function refSource(ref: WorkbookRef): string {
  if (ref.kind === 'range') {
    return formatRangeRef(ref.range)
  }
  if (ref.kind === 'name') {
    return ref.name
  }
  if (ref.kind === 'column' && ref.rows === undefined && ref.table.name) {
    return `${ref.table.name}[${ref.name}]`
  }
  return placeholderRefSource(ref)
}

function operandSource(operand: WorkbookFormulaOperand): string {
  if (isFormulaExpression(operand)) {
    return operand.source
  }
  if (isWorkbookRef(operand)) {
    return refSource(operand)
  }
  if (typeof operand === 'number') {
    if (!Number.isFinite(operand)) {
      throw new Error('Formula numbers must be finite')
    }
    return String(operand)
  }
  if (typeof operand === 'boolean') {
    return operand ? 'TRUE()' : 'FALSE()'
  }
  throw new Error('Formula operands must be formula expressions, workbook refs, finite numbers, booleans, or formula.text/raw wrappers')
}

function operandInputs(operand: WorkbookFormulaOperand): readonly WorkbookRef[] {
  if (isFormulaExpression(operand)) {
    return operand.inputs
  }
  if (isWorkbookRef(operand)) {
    return [operand]
  }
  return []
}

function operandLabels(operand: WorkbookFormulaOperand): readonly WorkbookFormulaLabel[] {
  if (isFormulaExpression(operand)) {
    return operand.labels
  }
  if (isWorkbookRef(operand)) {
    return [labelForRef(operand)]
  }
  return []
}

function collectInputs(args: readonly WorkbookFormulaOperand[]): readonly WorkbookRef[] {
  return uniqueRefs(args.flatMap((arg) => operandInputs(arg)))
}

function collectLabels(args: readonly WorkbookFormulaOperand[]): readonly WorkbookFormulaLabel[] {
  return uniqueLabels(args.flatMap((arg) => operandLabels(arg)))
}

function binary(left: WorkbookFormulaOperand, operator: '+' | '-' | '*' | '/', right: WorkbookFormulaOperand): WorkbookFormulaExpression {
  return createFormulaExpression(
    `(${operandSource(left)})${operator}(${operandSource(right)})`,
    collectInputs([left, right]),
    collectLabels([left, right]),
  )
}

function call(name: string, args: readonly WorkbookFormulaOperand[]): WorkbookFormulaExpression {
  const callee = normalizeFormulaFunctionName(name)
  return createFormulaExpression(`${callee}(${args.map(operandSource).join(',')})`, collectInputs(args), collectLabels(args))
}

export const formula = {
  raw(source: string, options: WorkbookRawFormulaOptions = {}): WorkbookFormulaExpression {
    const inputs = options.inputs ?? []
    return createFormulaExpression(source, inputs, options.labels ?? labelsForRefs(inputs))
  },
  source(expression: WorkbookFormulaOperand): string {
    return normalizeFormulaSource(operandSource(expression))
  },
  inputs(expression: WorkbookFormulaOperand): readonly WorkbookRef[] {
    return operandInputs(expression)
  },
  labels(expression: WorkbookFormulaOperand): readonly WorkbookFormulaLabel[] {
    return operandLabels(expression)
  },
  ref(ref: WorkbookRef): WorkbookFormulaExpression {
    return createFormulaExpression(refSource(ref), [ref], [labelForRef(ref)])
  },
  text(value: string): WorkbookFormulaExpression {
    return createFormulaExpression(`"${value.replaceAll('"', '""')}"`)
  },
  call,
  add(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '+', right)
  },
  subtract(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '-', right)
  },
  multiply(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '*', right)
  },
  divide(left: WorkbookFormulaOperand, right: WorkbookFormulaOperand): WorkbookFormulaExpression {
    return binary(left, '/', right)
  },
  sum(...args: readonly WorkbookFormulaOperand[]): WorkbookFormulaExpression {
    return call('SUM', args)
  },
} as const
