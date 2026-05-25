import { parseFormula, serializeFormula, type FormulaNode } from '@bilig/formula'

export interface WorkbookFormulaLabelReplacement {
  readonly name: string
  readonly source: string
}

type FormulaReferenceNode = Extract<
  FormulaNode,
  | { kind: 'NameRef' }
  | { kind: 'StructuredRef' }
  | { kind: 'CellRef' }
  | { kind: 'SpillRef' }
  | { kind: 'RowRef' }
  | { kind: 'ColumnRef' }
  | { kind: 'RangeRef' }
>

function isReferenceNode(node: FormulaNode): node is FormulaReferenceNode {
  return (
    node.kind === 'NameRef' ||
    node.kind === 'StructuredRef' ||
    node.kind === 'CellRef' ||
    node.kind === 'SpillRef' ||
    node.kind === 'RowRef' ||
    node.kind === 'ColumnRef' ||
    node.kind === 'RangeRef'
  )
}

function localName(node: FormulaNode | undefined): string | null {
  return node?.kind === 'NameRef' && node.sheetName === undefined ? node.name.toLocaleLowerCase('en-US') : null
}

function referenceKey(node: FormulaReferenceNode): string {
  return serializeFormula(node)
}

function formulaSource(source: string): string {
  const trimmed = source.trim()
  const normalized = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  if (normalized.length === 0) {
    throw new Error('Formula source cannot be empty')
  }
  return normalized
}

function parseReferenceLabel(labelName: string): FormulaReferenceNode | null {
  let labelAst: FormulaNode
  try {
    labelAst = parseFormula(formulaSource(labelName))
  } catch {
    return null
  }
  return isReferenceNode(labelAst) ? labelAst : null
}

function addReference(node: FormulaReferenceNode, localNames: ReadonlySet<string>, output: Set<string>): void {
  const name = localName(node)
  if (name !== null && localNames.has(name)) {
    return
  }
  output.add(referenceKey(node))
}

function collectCallReferences(
  node: Extract<FormulaNode, { kind: 'CallExpr' }>,
  localNames: ReadonlySet<string>,
  output: Set<string>,
): void {
  const callee = node.callee.toLocaleUpperCase('en-US')
  if (callee === 'LET') {
    collectLetReferences(node.args, localNames, output)
    return
  }
  if (callee === 'LAMBDA') {
    collectLambdaReferences(node.args, localNames, output)
    return
  }
  node.args.forEach((arg) => collectFormulaReferences(arg, localNames, output))
}

function collectLetReferences(args: readonly FormulaNode[], localNames: ReadonlySet<string>, output: Set<string>): void {
  const scoped = new Set(localNames)
  for (let index = 0; index < args.length - 2; index += 2) {
    const value = args[index + 1]
    if (value !== undefined) {
      collectFormulaReferences(value, scoped, output)
    }
    const name = localName(args[index])
    if (name !== null) {
      scoped.add(name)
    }
  }
  const body = args.at(-1)
  if (body !== undefined) {
    collectFormulaReferences(body, scoped, output)
  }
}

function collectLambdaReferences(args: readonly FormulaNode[], localNames: ReadonlySet<string>, output: Set<string>): void {
  const body = args.at(-1)
  if (body === undefined) {
    return
  }
  const scoped = new Set(localNames)
  args.slice(0, -1).forEach((arg) => {
    const name = localName(arg)
    if (name !== null) {
      scoped.add(name)
    }
  })
  collectFormulaReferences(body, scoped, output)
}

function collectFormulaReferences(node: FormulaNode, localNames: ReadonlySet<string>, output: Set<string>): void {
  if (isReferenceNode(node)) {
    addReference(node, localNames, output)
    return
  }

  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
      return
    case 'ArrayConstant':
      node.rows.forEach((row) => row.forEach((entry) => collectFormulaReferences(entry, localNames, output)))
      return
    case 'UnaryExpr':
      collectFormulaReferences(node.argument, localNames, output)
      return
    case 'BinaryExpr':
      collectFormulaReferences(node.left, localNames, output)
      collectFormulaReferences(node.right, localNames, output)
      return
    case 'CallExpr':
      collectCallReferences(node, localNames, output)
      return
    case 'InvokeExpr':
      collectFormulaReferences(node.callee, localNames, output)
      node.args.forEach((arg) => collectFormulaReferences(arg, localNames, output))
      return
  }
}

function replaceLetReferences(
  args: readonly FormulaNode[],
  localNames: ReadonlySet<string>,
  replacements: ReadonlyMap<string, FormulaNode>,
): FormulaNode[] {
  const scoped = new Set(localNames)
  const replaced: FormulaNode[] = []
  for (let index = 0; index < args.length - 1; index += 2) {
    const nameNode = args[index]
    if (nameNode !== undefined) {
      replaced.push(nameNode)
    }
    const valueNode = args[index + 1]
    if (valueNode !== undefined && index + 1 < args.length - 1) {
      replaced.push(replaceFormulaReferences(valueNode, scoped, replacements))
    }
    const name = localName(nameNode)
    if (name !== null) {
      scoped.add(name)
    }
  }
  const body = args.at(-1)
  if (body !== undefined) {
    replaced.push(replaceFormulaReferences(body, scoped, replacements))
  }
  return replaced
}

function replaceLambdaReferences(
  args: readonly FormulaNode[],
  localNames: ReadonlySet<string>,
  replacements: ReadonlyMap<string, FormulaNode>,
): FormulaNode[] {
  const body = args.at(-1)
  if (body === undefined) {
    return [...args]
  }
  const scoped = new Set(localNames)
  args.slice(0, -1).forEach((arg) => {
    const name = localName(arg)
    if (name !== null) {
      scoped.add(name)
    }
  })
  return [...args.slice(0, -1), replaceFormulaReferences(body, scoped, replacements)]
}

function replaceCallReferences(
  node: Extract<FormulaNode, { kind: 'CallExpr' }>,
  localNames: ReadonlySet<string>,
  replacements: ReadonlyMap<string, FormulaNode>,
): FormulaNode {
  const callee = node.callee.toLocaleUpperCase('en-US')
  if (callee === 'LET') {
    return { ...node, args: replaceLetReferences(node.args, localNames, replacements) }
  }
  if (callee === 'LAMBDA') {
    return { ...node, args: replaceLambdaReferences(node.args, localNames, replacements) }
  }
  return { ...node, args: node.args.map((arg) => replaceFormulaReferences(arg, localNames, replacements)) }
}

function replaceFormulaReferences(
  node: FormulaNode,
  localNames: ReadonlySet<string>,
  replacements: ReadonlyMap<string, FormulaNode>,
): FormulaNode {
  if (isReferenceNode(node)) {
    const name = localName(node)
    if (name !== null && localNames.has(name)) {
      return node
    }
    return replacements.get(referenceKey(node)) ?? node
  }

  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'OmittedArgument':
      return node
    case 'ArrayConstant':
      return { ...node, rows: node.rows.map((row) => row.map((entry) => replaceFormulaReferences(entry, localNames, replacements))) }
    case 'UnaryExpr':
      return { ...node, argument: replaceFormulaReferences(node.argument, localNames, replacements) }
    case 'BinaryExpr':
      return {
        ...node,
        left: replaceFormulaReferences(node.left, localNames, replacements),
        right: replaceFormulaReferences(node.right, localNames, replacements),
      }
    case 'CallExpr':
      return replaceCallReferences(node, localNames, replacements)
    case 'InvokeExpr':
      return {
        ...node,
        callee: replaceFormulaReferences(node.callee, localNames, replacements),
        args: node.args.map((arg) => replaceFormulaReferences(arg, localNames, replacements)),
      }
  }
}

export function formulaUsesLabel(ast: FormulaNode, labelName: string): boolean {
  const labelAst = parseReferenceLabel(labelName)
  if (labelAst === null) {
    return false
  }
  const references = new Set<string>()
  collectFormulaReferences(ast, new Set(), references)
  return references.has(referenceKey(labelAst))
}

export function materializeFormulaLabels(source: string, replacements: readonly WorkbookFormulaLabelReplacement[]): string {
  const sourceAst = parseFormula(formulaSource(source))
  const replacementMap = new Map<string, FormulaNode>()
  replacements.forEach((replacement) => {
    const labelAst = parseReferenceLabel(replacement.name)
    if (labelAst === null) {
      throw new Error(`Formula label ${replacement.name} must parse as a formula reference`)
    }
    const key = referenceKey(labelAst)
    const replacementAst = parseFormula(formulaSource(replacement.source))
    const previous = replacementMap.get(key)
    if (previous !== undefined && serializeFormula(previous) !== serializeFormula(replacementAst)) {
      throw new Error(`Formula label ${replacement.name} has conflicting replacements`)
    }
    replacementMap.set(key, replacementAst)
  })
  return serializeFormula(replaceFormulaReferences(sourceAst, new Set(), replacementMap))
}
