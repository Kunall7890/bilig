import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup-core-helpers.js'

interface LookupQueryBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  arrayResult: (values: CellValue[], rows: number, cols: number) => ReturnType<LookupBuiltin>
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toInteger: (value: CellValue | undefined) => number | undefined
  toStringValue: (value: CellValue | undefined) => string
  compareScalars: (left: CellValue, right: CellValue) => number | undefined
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue
}

interface QueryClauses {
  select: string
  where?: string
  orderBy?: string
  limit?: string
  offset?: string
}

interface QueryPlan {
  selectedCols: number[]
  where?: QueryCondition
  orderBy?: QueryOrder
  limit?: number
  offset?: number
}

interface QueryCondition {
  readonly col: number
  readonly op: '=' | '!=' | '<>' | '<' | '<=' | '>' | '>='
  readonly value: CellValue
}

interface QueryOrder {
  readonly col: number
  readonly direction: 'asc' | 'desc'
}

interface QueryRow {
  readonly sourceRow: number
  readonly values: CellValue[]
}

type QueryClauseKey = 'select' | 'where' | 'orderBy' | 'limit' | 'offset'

const clausePattern = /\b(select|where|order\s+by|limit|offset)\b/giu
const unsupportedClausePattern = /\b(group\s+by|pivot|having|label|format|options)\b/iu
const emptyValue: CellValue = { tag: ValueTag.Empty }

function normalizeQuery(query: string): string {
  return query.replace(/\s+/gu, ' ').trim().replace(/;$/u, '').trim()
}

function clauseKey(keyword: string): QueryClauseKey | undefined {
  const normalized = keyword.toLowerCase().replace(/\s+/gu, ' ')
  switch (normalized) {
    case 'select':
    case 'where':
    case 'limit':
    case 'offset':
      return normalized
    case 'order by':
      return 'orderBy'
    default:
      return undefined
  }
}

function parseClauses(query: string): QueryClauses | undefined {
  const normalized = normalizeQuery(query)
  if (normalized === '' || unsupportedClausePattern.test(normalized)) {
    return undefined
  }

  const matches = [...normalized.matchAll(clausePattern)]
  if (matches.length === 0 || matches[0]?.[1]?.toLowerCase() !== 'select') {
    return undefined
  }

  const clauses: Partial<QueryClauses> = {}
  for (let index = 0; index < matches.length; index += 1) {
    const match = matches[index]
    const keyword = match?.[1]
    if (!match || keyword === undefined || match.index === undefined) {
      return undefined
    }
    const key = clauseKey(keyword)
    if (key === undefined || clauses[key] !== undefined) {
      return undefined
    }
    const next = matches[index + 1]
    const start = match.index + keyword.length
    const end = next?.index ?? normalized.length
    const value = normalized.slice(start, end).trim()
    if (value === '') {
      return undefined
    }
    clauses[key] = value
  }

  if (clauses.select === undefined) {
    return undefined
  }
  return {
    select: clauses.select,
    ...(clauses.where === undefined ? {} : { where: clauses.where }),
    ...(clauses.orderBy === undefined ? {} : { orderBy: clauses.orderBy }),
    ...(clauses.limit === undefined ? {} : { limit: clauses.limit }),
    ...(clauses.offset === undefined ? {} : { offset: clauses.offset }),
  }
}

function columnIndexForRef(ref: string, cols: number): number | undefined {
  const colMatch = /^col([1-9]\d*)$/iu.exec(ref)
  if (colMatch) {
    const index = Number(colMatch[1]) - 1
    return index >= 0 && index < cols ? index : undefined
  }

  if (!/^[a-z]+$/iu.test(ref)) {
    return undefined
  }

  let index = 0
  for (const char of ref.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64
  }
  index -= 1
  return index >= 0 && index < cols ? index : undefined
}

function parseSelect(select: string, cols: number): number[] | undefined {
  if (select.trim() === '*') {
    return Array.from({ length: cols }, (_value, index) => index)
  }

  const selected = select.split(',').map((part) => part.trim())
  if (selected.length === 0 || selected.some((part) => part === '')) {
    return undefined
  }

  const indexes: number[] = []
  for (const part of selected) {
    const col = columnIndexForRef(part, cols)
    if (col === undefined) {
      return undefined
    }
    indexes.push(col)
  }
  return indexes
}

function textValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function parseLiteral(raw: string): CellValue | undefined {
  const value = raw.trim()
  const quoted = /^'(.*)'$/u.exec(value)
  if (quoted) {
    return textValue((quoted[1] ?? '').replace(/''/gu, "'"))
  }

  if (/^(true|false)$/iu.test(value)) {
    return { tag: ValueTag.Boolean, value: /^true$/iu.test(value) }
  }

  if (value !== '') {
    const numeric = Number(value)
    if (Number.isFinite(numeric)) {
      return { tag: ValueTag.Number, value: numeric }
    }
  }

  return undefined
}

function parseWhere(where: string, cols: number): QueryCondition | undefined {
  const match = /^([a-z]+|col[1-9]\d*)\s*(=|!=|<>|<=|>=|<|>)\s*(.+)$/iu.exec(where.trim())
  if (!match) {
    return undefined
  }
  const col = columnIndexForRef(match[1] ?? '', cols)
  const value = parseLiteral(match[3] ?? '')
  if (col === undefined || value === undefined) {
    return undefined
  }
  const op = parseOperator(match[2] ?? '')
  if (op === undefined) {
    return undefined
  }
  return {
    col,
    op,
    value,
  }
}

function parseOperator(raw: string): QueryCondition['op'] | undefined {
  switch (raw) {
    case '=':
    case '!=':
    case '<>':
    case '<':
    case '<=':
    case '>':
    case '>=':
      return raw
    default:
      return undefined
  }
}

function parseDirection(raw: string | undefined): QueryOrder['direction'] | undefined {
  if (raw === undefined) {
    return 'asc'
  }
  const normalized = raw.toLowerCase()
  if (normalized === 'asc' || normalized === 'desc') {
    return normalized
  }
  return undefined
}

function parseOrder(orderBy: string, cols: number): QueryOrder | undefined {
  const match = /^([a-z]+|col[1-9]\d*)(?:\s+(asc|desc))?$/iu.exec(orderBy.trim())
  if (!match) {
    return undefined
  }
  const col = columnIndexForRef(match[1] ?? '', cols)
  const direction = parseDirection(match[2])
  if (col === undefined || direction === undefined) {
    return undefined
  }
  return {
    col,
    direction,
  }
}

function parseNonNegativeInteger(raw: string | undefined): number | undefined {
  if (raw === undefined || !/^\d+$/u.test(raw.trim())) {
    return undefined
  }
  return Number(raw.trim())
}

function parsePlan(query: string, cols: number): QueryPlan | undefined {
  const clauses = parseClauses(query)
  if (!clauses) {
    return undefined
  }

  const selectedCols = parseSelect(clauses.select, cols)
  if (!selectedCols) {
    return undefined
  }

  const where = clauses.where === undefined ? undefined : parseWhere(clauses.where, cols)
  if (clauses.where !== undefined && where === undefined) {
    return undefined
  }

  const orderBy = clauses.orderBy === undefined ? undefined : parseOrder(clauses.orderBy, cols)
  if (clauses.orderBy !== undefined && orderBy === undefined) {
    return undefined
  }

  const limit = clauses.limit === undefined ? undefined : parseNonNegativeInteger(clauses.limit)
  if (clauses.limit !== undefined && limit === undefined) {
    return undefined
  }

  const offset = clauses.offset === undefined ? undefined : parseNonNegativeInteger(clauses.offset)
  if (clauses.offset !== undefined && offset === undefined) {
    return undefined
  }

  const plan: QueryPlan = { selectedCols }
  if (where !== undefined) {
    plan.where = where
  }
  if (orderBy !== undefined) {
    plan.orderBy = orderBy
  }
  if (limit !== undefined) {
    plan.limit = limit
  }
  if (offset !== undefined) {
    plan.offset = offset
  }
  return plan
}

function compareWithOperator(deps: LookupQueryBuiltinDeps, left: CellValue, condition: QueryCondition): boolean | CellValue {
  if (left.tag === ValueTag.Error) {
    return left
  }
  const comparison = deps.compareScalars(left, condition.value)
  if (comparison === undefined) {
    return condition.op === '!=' || condition.op === '<>'
  }
  switch (condition.op) {
    case '=':
      return comparison === 0
    case '!=':
    case '<>':
      return comparison !== 0
    case '<':
      return comparison < 0
    case '<=':
      return comparison <= 0
    case '>':
      return comparison > 0
    case '>=':
      return comparison >= 0
  }
}

function projectRows(rows: readonly QueryRow[], selectedCols: readonly number[]): CellValue[] {
  const values: CellValue[] = []
  for (const row of rows) {
    for (const col of selectedCols) {
      values.push(row.values[col] ?? emptyValue)
    }
  }
  return values
}

function createRows(deps: LookupQueryBuiltinDeps, range: RangeBuiltinArgument, startRow: number): QueryRow[] {
  const rows: QueryRow[] = []
  for (let row = startRow; row < range.rows; row += 1) {
    const values: CellValue[] = []
    for (let col = 0; col < range.cols; col += 1) {
      values.push(deps.getRangeValue(range, row, col))
    }
    rows.push({ sourceRow: row, values })
  }
  return rows
}

function queryRange(
  deps: LookupQueryBuiltinDeps,
  range: RangeBuiltinArgument,
  query: string,
  headerRows: number,
): ReturnType<LookupBuiltin> {
  if (range.rows < 0 || range.cols <= 0 || headerRows > range.rows) {
    return deps.errorValue(ErrorCode.Value)
  }

  const plan = parsePlan(query, range.cols)
  if (!plan) {
    return deps.errorValue(ErrorCode.Value)
  }

  const header = headerRows > 0 ? createRows(deps, range, 0).slice(0, headerRows) : []
  let body = createRows(deps, range, headerRows)

  if (plan.where) {
    const kept: QueryRow[] = []
    for (const row of body) {
      const matches = compareWithOperator(deps, row.values[plan.where.col] ?? emptyValue, plan.where)
      if (typeof matches !== 'boolean') {
        return matches
      }
      if (matches) {
        kept.push(row)
      }
    }
    body = kept
  }

  if (plan.orderBy) {
    const order = plan.orderBy
    let sortError: CellValue | undefined
    body = body.toSorted((left, right) => {
      const leftValue = left.values[order.col] ?? emptyValue
      const rightValue = right.values[order.col] ?? emptyValue
      if (leftValue.tag === ValueTag.Error) {
        sortError = leftValue
        return 0
      }
      if (rightValue.tag === ValueTag.Error) {
        sortError = rightValue
        return 0
      }
      const comparison = deps.compareScalars(leftValue, rightValue)
      if (comparison === undefined) {
        sortError = deps.errorValue(ErrorCode.Value)
        return 0
      }
      if (comparison === 0) {
        return left.sourceRow - right.sourceRow
      }
      return order.direction === 'desc' ? -comparison : comparison
    })
    if (sortError) {
      return sortError
    }
  }

  const offset = plan.offset ?? 0
  const limitedBody = plan.limit === undefined ? body.slice(offset) : body.slice(offset, offset + plan.limit)
  const outputRows = [...header, ...limitedBody]
  return deps.arrayResult(projectRows(outputRows, plan.selectedCols), outputRows.length, plan.selectedCols.length)
}

export function createLookupQueryBuiltins(deps: LookupQueryBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    QUERY: (dataArg, queryArg, headersArg) => {
      if (deps.isError(dataArg)) {
        return dataArg
      }
      if (deps.isError(queryArg)) {
        return queryArg
      }
      if (deps.isError(headersArg)) {
        return headersArg
      }

      const data = deps.requireCellRange(dataArg)
      if (!deps.isRangeArg(data) || deps.isRangeArg(queryArg) || deps.isRangeArg(headersArg) || queryArg === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }

      const query = deps.toStringValue(queryArg)
      const headerRows = headersArg === undefined ? 0 : deps.toInteger(headersArg)
      if (headerRows === undefined || headerRows < 0) {
        return deps.errorValue(ErrorCode.Value)
      }

      return queryRange(deps, data, query, headerRows)
    },
  }
}
