import { compileFormula, normalizeFormulaFunctionName, parseFormula, type ParsedDependencyReference } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import { isWorkbookRef, type WorkbookRef } from './find.js'

export interface WorkbookFormulaExpression {
  readonly kind: 'formula'
  readonly source: string
  readonly inputs: readonly WorkbookRef[]
}

export interface WorkbookRawFormulaOptions {
  readonly inputs?: readonly WorkbookRef[]
}

export type WorkbookFormulaDependency =
  | {
      readonly kind: 'cell'
      readonly address: string
      readonly sheetName?: string
      readonly explicitSheet?: boolean
      readonly row?: number
      readonly col?: number
      readonly rowAbsolute?: boolean
      readonly colAbsolute?: boolean
    }
  | {
      readonly kind: 'range'
      readonly address: string
      readonly refKind: 'cells' | 'rows' | 'cols'
      readonly sheetName?: string
      readonly sheetEndName?: string
      readonly explicitSheet?: boolean
      readonly startAddress: string
      readonly endAddress: string
      readonly startRow: number
      readonly endRow: number
      readonly startCol: number
      readonly endCol: number
      readonly startRowAbsolute?: boolean
      readonly endRowAbsolute?: boolean
      readonly startColAbsolute?: boolean
      readonly endColAbsolute?: boolean
    }

export interface WorkbookFormulaInspection {
  readonly source: string
  readonly inputs: readonly WorkbookRef[]
  readonly dependencies: readonly WorkbookFormulaDependency[]
  readonly names: readonly string[]
  readonly tables: readonly string[]
  readonly spills: readonly string[]
  readonly volatile: boolean
  readonly producesSpill: boolean
}

export type WorkbookFormulaOperand = WorkbookFormulaExpression | WorkbookRef | string | number | boolean

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
  if (!Array.isArray(refs)) {
    throw new Error('Formula inputs must be an array')
  }
  const seen = new Set<string>()
  const unique: WorkbookRef[] = []
  refs.forEach((ref, index) => {
    if (!isWorkbookRef(ref)) {
      throw new Error(`Formula input at inputs[${index.toString()}] must be a WorkbookRef`)
    }
    const key = `${ref.kind}:${ref.id}`
    if (seen.has(key)) {
      return
    }
    seen.add(key)
    unique.push(ref)
  })
  return Object.freeze(unique)
}

function createFormulaExpression(source: string, inputs: readonly WorkbookRef[] = []): WorkbookFormulaExpression {
  return Object.freeze({
    kind: 'formula',
    source: normalizeFormulaSource(source),
    inputs: uniqueRefs(inputs),
  })
}

function isFormulaExpression(value: unknown): value is WorkbookFormulaExpression {
  return (
    typeof value === 'object' &&
    value !== null &&
    'kind' in value &&
    value.kind === 'formula' &&
    'source' in value &&
    typeof value.source === 'string' &&
    'inputs' in value &&
    Array.isArray(value.inputs) &&
    value.inputs.every(isWorkbookRef)
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
  if (typeof operand === 'string') {
    return operand
  }
  throw new Error('Formula operand must be a formula expression, WorkbookRef, string, finite number, or boolean')
}

function operandInputs(operand: WorkbookFormulaOperand): readonly WorkbookRef[] {
  if (isFormulaExpression(operand)) {
    return operand.inputs
  }
  if (isWorkbookRef(operand)) {
    return [operand]
  }
  if (typeof operand === 'object' && operand !== null) {
    throw new Error('Formula operand must be a formula expression, WorkbookRef, string, finite number, or boolean')
  }
  return []
}

function collectInputs(args: readonly WorkbookFormulaOperand[]): readonly WorkbookRef[] {
  return uniqueRefs(args.flatMap((arg) => operandInputs(arg)))
}

function cloneFormulaDependency(dependency: ParsedDependencyReference): WorkbookFormulaDependency {
  if (dependency.kind === 'cell') {
    return Object.freeze({
      kind: 'cell',
      address: dependency.address,
      ...(dependency.sheetName !== undefined ? { sheetName: dependency.sheetName } : {}),
      ...(dependency.explicitSheet !== undefined ? { explicitSheet: dependency.explicitSheet } : {}),
      ...(dependency.row !== undefined ? { row: dependency.row } : {}),
      ...(dependency.col !== undefined ? { col: dependency.col } : {}),
      ...(dependency.rowAbsolute !== undefined ? { rowAbsolute: dependency.rowAbsolute } : {}),
      ...(dependency.colAbsolute !== undefined ? { colAbsolute: dependency.colAbsolute } : {}),
    })
  }
  return Object.freeze({
    kind: 'range',
    address: dependency.address,
    refKind: dependency.refKind,
    ...(dependency.sheetName !== undefined ? { sheetName: dependency.sheetName } : {}),
    ...(dependency.sheetEndName !== undefined ? { sheetEndName: dependency.sheetEndName } : {}),
    ...(dependency.explicitSheet !== undefined ? { explicitSheet: dependency.explicitSheet } : {}),
    startAddress: dependency.startAddress,
    endAddress: dependency.endAddress,
    startRow: dependency.startRow,
    endRow: dependency.endRow,
    startCol: dependency.startCol,
    endCol: dependency.endCol,
    ...(dependency.startRowAbsolute !== undefined ? { startRowAbsolute: dependency.startRowAbsolute } : {}),
    ...(dependency.endRowAbsolute !== undefined ? { endRowAbsolute: dependency.endRowAbsolute } : {}),
    ...(dependency.startColAbsolute !== undefined ? { startColAbsolute: dependency.startColAbsolute } : {}),
    ...(dependency.endColAbsolute !== undefined ? { endColAbsolute: dependency.endColAbsolute } : {}),
  })
}

function inspectFormula(expression: WorkbookFormulaOperand): WorkbookFormulaInspection {
  const source = normalizeFormulaSource(operandSource(expression))
  const compiled = compileFormula(source)
  return Object.freeze({
    source,
    inputs: Object.freeze([...operandInputs(expression)]),
    dependencies: Object.freeze((compiled.parsedDeps ?? []).map(cloneFormulaDependency)),
    names: Object.freeze([...compiled.symbolicNames]),
    tables: Object.freeze([...compiled.symbolicTables]),
    spills: Object.freeze([...compiled.symbolicSpills]),
    volatile: compiled.volatile,
    producesSpill: compiled.producesSpill,
  })
}

function binary(left: WorkbookFormulaOperand, operator: '+' | '-' | '*' | '/', right: WorkbookFormulaOperand): WorkbookFormulaExpression {
  return createFormulaExpression(`(${operandSource(left)})${operator}(${operandSource(right)})`, collectInputs([left, right]))
}

function call(name: string, args: readonly WorkbookFormulaOperand[]): WorkbookFormulaExpression {
  const callee = normalizeFormulaFunctionName(name)
  return createFormulaExpression(`${callee}(${args.map(operandSource).join(',')})`, collectInputs(args))
}

export const formula = {
  raw(source: string, options: WorkbookRawFormulaOptions = {}): WorkbookFormulaExpression {
    return createFormulaExpression(source, options.inputs ?? [])
  },
  source(expression: WorkbookFormulaOperand): string {
    return normalizeFormulaSource(operandSource(expression))
  },
  inputs(expression: WorkbookFormulaOperand): readonly WorkbookRef[] {
    return Object.freeze([...operandInputs(expression)])
  },
  dependencies(expression: WorkbookFormulaOperand): readonly WorkbookFormulaDependency[] {
    return inspectFormula(expression).dependencies
  },
  inspect(expression: WorkbookFormulaOperand): WorkbookFormulaInspection {
    return inspectFormula(expression)
  },
  ref(ref: WorkbookRef): WorkbookFormulaExpression {
    return createFormulaExpression(refSource(ref), [ref])
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
