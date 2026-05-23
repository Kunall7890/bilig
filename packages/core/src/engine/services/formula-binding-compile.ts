import { compileFormulaAst, serializeFormula } from '@bilig/formula'
import type { FormulaNode } from '@bilig/formula'
import { resolveMetadataReferencesInAst, type MetadataFormulaValueContext } from '../../engine-metadata-utils.js'
import type { FormulaTemplateResolution } from '../../formula/template-bank.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export function compileFormulaBindingForCell(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly cellIndex: number
  readonly currentSheetName: string
  readonly source: string
  readonly resolvedCompiledCache: Map<string, ParsedCompiledFormula>
  readonly normalizeLookupCompileMode: (compiled: ParsedCompiledFormula) => ParsedCompiledFormula
}): { compiled: ParsedCompiledFormula; templateResolution: FormulaTemplateResolution } {
  const position = args.serviceArgs.state.workbook.getCellPosition(args.cellIndex)
  if (!position) {
    throw new Error(`Cannot resolve formula template without coordinates for cell ${args.cellIndex}`)
  }
  const templateResolution = args.serviceArgs.resolveTemplateForCell(args.source, position.row, position.col)
  const compiled = args.normalizeLookupCompileMode(templateResolution.compiled as ParsedCompiledFormula)
  if (
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    !hasUnqualifiedStructuredReference(compiled.ast)
  ) {
    return {
      compiled,
      templateResolution,
    }
  }

  const resolved = resolveMetadataReferencesInAst(
    compiled.ast,
    {
      resolveName: (name, scopeSheetName) =>
        args.serviceArgs.state.workbook.getDefinedName(name, scopeSheetName ?? args.currentSheetName)?.value,
      resolveStructuredReference: (tableName, columnName, options) =>
        args.serviceArgs.resolveStructuredReference(tableName, columnName, {
          ...options,
          ownerSheetName: args.currentSheetName,
          ownerAddress: args.serviceArgs.state.workbook.getAddress(args.cellIndex),
        }),
      resolveSpillReference: (sheetName, address) => args.serviceArgs.resolveSpillReference(args.currentSheetName, sheetName, address),
    },
    new Set<string>(),
    metadataValueContextForCell(args.serviceArgs.state.workbook, args.currentSheetName, args.cellIndex),
  )
  if (!resolved.substituted || !resolved.fullyResolved) {
    return {
      compiled,
      templateResolution,
    }
  }

  const resolvedCacheKey = `${args.currentSheetName}\u0000${args.source}\u0000${serializeFormula(resolved.node)}`
  let resolvedCompiled = args.resolvedCompiledCache.get(resolvedCacheKey)
  if (!resolvedCompiled) {
    resolvedCompiled = args.normalizeLookupCompileMode(
      compileFormulaAst(args.source, resolved.node, {
        originalAst: compiled.ast,
        symbolicNames: compiled.symbolicNames,
        symbolicTables: compiled.symbolicTables.filter((tableName) => tableName.length > 0),
        symbolicSpills: compiled.symbolicSpills,
      }) as ParsedCompiledFormula,
    )
    args.resolvedCompiledCache.set(resolvedCacheKey, resolvedCompiled)
  }
  return {
    compiled: resolvedCompiled,
    templateResolution,
  }
}

function hasUnqualifiedStructuredReference(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'StructuredRef':
      return node.tableName.length === 0
    case 'ArrayConstant':
      return node.rows.some((row) => row.some(hasUnqualifiedStructuredReference))
    case 'UnaryExpr':
      return hasUnqualifiedStructuredReference(node.argument)
    case 'BinaryExpr':
      return hasUnqualifiedStructuredReference(node.left) || hasUnqualifiedStructuredReference(node.right)
    case 'CallExpr':
      return node.args.some(hasUnqualifiedStructuredReference)
    case 'InvokeExpr':
      return hasUnqualifiedStructuredReference(node.callee) || node.args.some(hasUnqualifiedStructuredReference)
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return false
  }
}

function metadataValueContextForCell(
  workbook: CreateEngineFormulaBindingServiceArgs['state']['workbook'],
  sheetName: string,
  cellIndex: number,
): MetadataFormulaValueContext {
  return workbook.getSpill(sheetName, workbook.getAddress(cellIndex)) ? 'array' : 'scalar'
}
