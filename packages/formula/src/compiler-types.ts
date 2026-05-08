import type { FormulaRecord } from '@bilig/protocol'
import type { FormulaNode } from './ast.js'
import type { JsPlanInstruction } from './js-evaluator.js'

export interface CompiledFormula extends FormulaRecord {
  ast: FormulaNode
  optimizedAst: FormulaNode
  astMatchesSource?: boolean
  directAggregateCandidate?: DirectAggregateCandidate
  deps: string[]
  parsedDeps?: ParsedDependencyReference[]
  symbolicNames: string[]
  symbolicTables: string[]
  symbolicSpills: string[]
  volatile: boolean
  randCallCount: number
  producesSpill: boolean
  jsPlan: JsPlanInstruction[]
  program: Uint32Array
  constants: Float64Array
  symbolicRefs: string[]
  parsedSymbolicRefs?: ParsedCellReferenceInfo[]
  symbolicRanges: string[]
  parsedSymbolicRanges?: ParsedRangeReferenceInfo[]
  symbolicStrings: string[]
}

export interface ParsedCellReferenceInfo {
  address: string
  sheetName?: string
  explicitSheet?: boolean
  row?: number
  col?: number
  rowAbsolute?: boolean
  colAbsolute?: boolean
}

export interface ParsedRangeReferenceInfo {
  address: string
  kind: 'range'
  refKind: 'cells' | 'rows' | 'cols'
  sheetName?: string
  sheetEndName?: string
  explicitSheet?: boolean
  startAddress: string
  endAddress: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
  startRowAbsolute?: boolean
  endRowAbsolute?: boolean
  startColAbsolute?: boolean
  endColAbsolute?: boolean
}

export type ParsedDependencyReference = ({ kind: 'cell' } & ParsedCellReferenceInfo) | ParsedRangeReferenceInfo

export interface DirectAggregateCandidate {
  callee: string
  aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
  symbolicRangeIndex: number
  resultOffset?: number
}

export interface CompileFormulaAstOptions {
  originalAst?: FormulaNode
  symbolicNames?: string[]
  symbolicTables?: string[]
  symbolicSpills?: string[]
}
