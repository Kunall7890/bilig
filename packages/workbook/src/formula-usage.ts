import { parseFormula, serializeFormula, type FormulaNode } from '@bilig/formula'

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

export function formulaUsesLabel(ast: FormulaNode, labelName: string): boolean {
  let labelAst: FormulaNode
  try {
    labelAst = parseFormula(labelName)
  } catch {
    return false
  }
  if (!isReferenceNode(labelAst)) {
    return false
  }
  const references = new Set<string>()
  collectFormulaReferences(ast, new Set(), references)
  return references.has(referenceKey(labelAst))
}
