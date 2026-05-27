import { hasBuiltin, parseFormula, type FormulaNode } from '@bilig/formula'
import type { WorkbookSnapshot } from '@bilig/protocol'

const UNSUPPORTED_FORMULA_CACHE_KEY_SEPARATOR = '\t'

function normalizeFormulaName(name: string): string {
  return name.trim().toUpperCase()
}

function normalizeFormulaCacheAddress(address: string): string {
  return address.trim().toUpperCase()
}

function unsupportedFormulaCacheKey(sheetName: string, address: string, formula: string): string {
  return [sheetName, normalizeFormulaCacheAddress(address), formula].join(UNSUPPORTED_FORMULA_CACHE_KEY_SEPARATOR)
}

type UnavailableCallMatcher = (normalizedName: string) => boolean

const matchAnyUnavailableFormulaCall: UnavailableCallMatcher = () => true

function isFullRecalcPreservableUnavailableFormulaCall(normalizedName: string): boolean {
  return normalizedName === '_FV'
}

function containsAsciiCaseInsensitive(source: string, needle: string): boolean {
  const maxStart = source.length - needle.length
  for (let start = source.indexOf('_'); start >= 0 && start <= maxStart; start = source.indexOf('_', start + 1)) {
    let matches = true
    for (let offset = 0; offset < needle.length; offset += 1) {
      const sourceCode = source.charCodeAt(start + offset)
      const needleCode = needle.charCodeAt(offset)
      if (sourceCode === needleCode) {
        continue
      }
      const foldedSourceCode = sourceCode >= 97 && sourceCode <= 122 ? sourceCode - 32 : sourceCode
      if (foldedSourceCode !== needleCode) {
        matches = false
        break
      }
    }
    if (matches) {
      return true
    }
  }
  return false
}

export function formulaMayContainFullRecalcPreservableUnavailableFormulaCall(source: string): boolean {
  return source.indexOf('_') !== -1 && containsAsciiCaseInsensitive(source, '_FV')
}

function withLocalFormulaName(localNames: ReadonlySet<string>, name: string): ReadonlySet<string> {
  const normalized = normalizeFormulaName(name)
  if (normalized.length === 0 || localNames.has(normalized)) {
    return localNames
  }
  const next = new Set(localNames)
  next.add(normalized)
  return next
}

export function collectDefinedFormulaNames(snapshot: WorkbookSnapshot): ReadonlySet<string> {
  const definedNames = snapshot.workbook.metadata?.definedNames
  if (!definedNames || definedNames.length === 0) {
    return new Set()
  }
  const names = new Set<string>()
  for (const definedName of definedNames) {
    const normalized = normalizeFormulaName(definedName.name)
    if (normalized.length > 0) {
      names.add(normalized)
    }
  }
  return names
}

export function collectPreservedUnsupportedFormulaCacheKeys(snapshot: WorkbookSnapshot): ReadonlySet<string> {
  const unsupportedFormulaDependencies = snapshot.workbook.metadata?.unsupportedFormulaDependencies
  if (!unsupportedFormulaDependencies || unsupportedFormulaDependencies.length === 0) {
    return new Set()
  }
  const keys = new Set<string>()
  for (const dependency of unsupportedFormulaDependencies) {
    if (dependency.cachedFormulaValuePreserved) {
      keys.add(unsupportedFormulaCacheKey(dependency.sheetName, dependency.address, dependency.importedFormula))
    }
  }
  return keys
}

export function formulaHasPreservedUnsupportedDependencyCache(
  preservedCacheKeys: ReadonlySet<string>,
  sheetName: string,
  address: string,
  formula: string,
): boolean {
  return preservedCacheKeys.has(unsupportedFormulaCacheKey(sheetName, address, formula))
}

function isAvailableFormulaCall(callee: string, definedNames: ReadonlySet<string>, localNames: ReadonlySet<string>): boolean {
  const normalized = normalizeFormulaName(callee)
  return normalized.length > 0 && (hasBuiltin(normalized) || definedNames.has(normalized) || localNames.has(normalized))
}

function lambdaBodyHasUnavailableCall(
  args: readonly FormulaNode[],
  definedNames: ReadonlySet<string>,
  localNames: ReadonlySet<string>,
  matcher: UnavailableCallMatcher,
): boolean {
  if (args.length === 0) {
    return false
  }
  let lambdaLocals = localNames
  for (let index = 0; index < args.length - 1; index += 1) {
    const param = args[index]
    if (param?.kind === 'NameRef') {
      lambdaLocals = withLocalFormulaName(lambdaLocals, param.name)
    }
  }
  return formulaNodeHasUnavailableCall(args[args.length - 1]!, definedNames, lambdaLocals, matcher)
}

function letBodyHasUnavailableCall(
  args: readonly FormulaNode[],
  definedNames: ReadonlySet<string>,
  localNames: ReadonlySet<string>,
  matcher: UnavailableCallMatcher,
): boolean {
  if (args.length < 2) {
    return args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames, matcher))
  }
  let letLocals = localNames
  const finalArgIndex = args.length - 1
  for (let index = 0; index < finalArgIndex; index += 2) {
    const valueNode = args[index + 1]
    if (valueNode && formulaNodeHasUnavailableCall(valueNode, definedNames, letLocals, matcher)) {
      return true
    }
    const nameNode = args[index]
    if (nameNode?.kind === 'NameRef') {
      letLocals = withLocalFormulaName(letLocals, nameNode.name)
    }
  }
  return formulaNodeHasUnavailableCall(args[finalArgIndex]!, definedNames, letLocals, matcher)
}

function formulaNodeHasUnavailableCall(
  node: FormulaNode,
  definedNames: ReadonlySet<string>,
  localNames: ReadonlySet<string>,
  matcher: UnavailableCallMatcher,
): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
    case 'NameRef':
    case 'StructuredRef':
    case 'CellRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
      return false
    case 'ArrayConstant':
      return node.rows.some((row) => row.some((entry) => formulaNodeHasUnavailableCall(entry, definedNames, localNames, matcher)))
    case 'UnaryExpr':
      return formulaNodeHasUnavailableCall(node.argument, definedNames, localNames, matcher)
    case 'BinaryExpr':
      return (
        formulaNodeHasUnavailableCall(node.left, definedNames, localNames, matcher) ||
        formulaNodeHasUnavailableCall(node.right, definedNames, localNames, matcher)
      )
    case 'CallExpr': {
      const normalized = normalizeFormulaName(node.callee)
      if (normalized === 'LAMBDA') {
        return lambdaBodyHasUnavailableCall(node.args, definedNames, localNames, matcher)
      }
      if (normalized === 'LET') {
        return letBodyHasUnavailableCall(node.args, definedNames, localNames, matcher)
      }
      if (!isAvailableFormulaCall(node.callee, definedNames, localNames)) {
        return matcher(normalized)
      }
      return node.args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames, matcher))
    }
    case 'InvokeExpr':
      return (
        formulaNodeHasUnavailableCall(node.callee, definedNames, localNames, matcher) ||
        node.args.some((arg) => formulaNodeHasUnavailableCall(arg, definedNames, localNames, matcher))
      )
  }
}

export function formulaShouldUseCachedUnsupportedFunctionValue(source: string, definedNames: ReadonlySet<string>): boolean {
  try {
    return formulaNodeHasUnavailableCall(parseFormula(source), definedNames, new Set(), matchAnyUnavailableFormulaCall)
  } catch {
    return false
  }
}

export function formulaShouldPreserveCachedUnsupportedFunctionValueOnFullRecalc(
  source: string,
  definedNames: ReadonlySet<string>,
): boolean {
  if (!formulaMayContainFullRecalcPreservableUnavailableFormulaCall(source)) {
    return false
  }
  try {
    return formulaNodeHasUnavailableCall(parseFormula(source), definedNames, new Set(), isFullRecalcPreservableUnavailableFormulaCall)
  } catch {
    return false
  }
}
