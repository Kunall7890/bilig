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
  groupBy?: string
  orderBy?: string
  limit?: string
  offset?: string
}

interface QueryPlan {
  select: QuerySelectItem[]
  where?: QueryCondition
  groupBy?: number[]
  orderBy?: QueryOrder
  limit?: number
  offset?: number
}

type QueryAggregateFunction = 'sum' | 'count'

interface QueryColumnSelectItem {
  readonly kind: 'column'
  readonly col: number
}

interface QueryAggregateSelectItem {
  readonly kind: 'aggregate'
  readonly fn: QueryAggregateFunction
  readonly col: number
}

type QuerySelectItem = QueryColumnSelectItem | QueryAggregateSelectItem

interface QueryCondition {
  readonly col: number
  readonly op: '=' | '!=' | '<>' | '<' | '<=' | '>' | '>='
  readonly value: CellValue
}

interface QueryOrder {
  readonly item: QuerySelectItem
  readonly direction: 'asc' | 'desc'
}

interface QueryRow {
  readonly sourceRow: number
  readonly values: CellValue[]
}

interface QueryAggregateOutputRow {
  readonly sourceRow: number
  readonly values: CellValue[]
  readonly groupRows: readonly QueryRow[]
}

type QueryClauseKey = 'select' | 'where' | 'groupBy' | 'orderBy' | 'limit' | 'offset'

const clausePattern = /\b(select|where|group\s+by|order\s+by|limit|offset)\b/giu
const unsupportedClausePattern = /\b(pivot|having|label|format|options)\b/iu
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
    case 'group by':
      return 'groupBy'
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
    ...(clauses.groupBy === undefined ? {} : { groupBy: clauses.groupBy }),
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

function parseAggregateFunction(raw: string | undefined): QueryAggregateFunction | undefined {
  const normalized = raw?.toLowerCase()
  if (normalized === 'sum' || normalized === 'count') {
    return normalized
  }
  return undefined
}

function parseSelectItem(part: string, cols: number): QuerySelectItem | undefined {
  const aggregateMatch = /^(sum|count)\(\s*([a-z]+|col[1-9]\d*)\s*\)$/iu.exec(part)
  if (aggregateMatch) {
    const col = columnIndexForRef(aggregateMatch[2] ?? '', cols)
    const fn = parseAggregateFunction(aggregateMatch[1])
    if (col === undefined || fn === undefined) {
      return undefined
    }
    return {
      kind: 'aggregate',
      fn,
      col,
    }
  }

  const col = columnIndexForRef(part, cols)
  return col === undefined ? undefined : { kind: 'column', col }
}

function parseSelect(select: string, cols: number): QuerySelectItem[] | undefined {
  if (select.trim() === '*') {
    return Array.from({ length: cols }, (_value, col) => ({ kind: 'column', col }))
  }

  const selected = select.split(',').map((part) => part.trim())
  if (selected.length === 0 || selected.some((part) => part === '')) {
    return undefined
  }

  const items: QuerySelectItem[] = []
  for (const part of selected) {
    const item = parseSelectItem(part, cols)
    if (item === undefined) {
      return undefined
    }
    items.push(item)
  }
  return items
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

function parseColumnList(raw: string, cols: number): number[] | undefined {
  const parts = raw.split(',').map((part) => part.trim())
  if (parts.length === 0 || parts.some((part) => part === '')) {
    return undefined
  }
  const columns: number[] = []
  for (const part of parts) {
    const col = columnIndexForRef(part, cols)
    if (col === undefined) {
      return undefined
    }
    columns.push(col)
  }
  return columns
}

function parseOrder(orderBy: string, cols: number): QueryOrder | undefined {
  const match = /^((?:sum|count)\(\s*(?:[a-z]+|col[1-9]\d*)\s*\)|[a-z]+|col[1-9]\d*)(?:\s+(asc|desc))?$/iu.exec(orderBy.trim())
  if (!match) {
    return undefined
  }
  const item = parseSelectItem(match[1] ?? '', cols)
  const direction = parseDirection(match[2])
  if (item === undefined || direction === undefined) {
    return undefined
  }
  return {
    item,
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

  const select = parseSelect(clauses.select, cols)
  if (!select) {
    return undefined
  }

  const groupBy = clauses.groupBy === undefined ? undefined : parseColumnList(clauses.groupBy, cols)
  if (clauses.groupBy !== undefined && groupBy === undefined) {
    return undefined
  }

  const hasAggregate = select.some((item) => item.kind === 'aggregate')
  if (groupBy !== undefined) {
    if (!hasAggregate) {
      return undefined
    }
    const groupedCols = new Set(groupBy)
    if (select.some((item) => item.kind === 'column' && !groupedCols.has(item.col))) {
      return undefined
    }
  }

  const where = clauses.where === undefined ? undefined : parseWhere(clauses.where, cols)
  if (clauses.where !== undefined && where === undefined) {
    return undefined
  }

  const orderBy = clauses.orderBy === undefined ? undefined : parseOrder(clauses.orderBy, cols)
  if (clauses.orderBy !== undefined && orderBy === undefined) {
    return undefined
  }

  if (orderBy?.item.kind === 'aggregate' && !hasAggregate) {
    return undefined
  }
  if (groupBy !== undefined && orderBy?.item.kind === 'column' && !new Set(groupBy).has(orderBy.item.col)) {
    return undefined
  }
  if (hasAggregate && groupBy === undefined && select.some((item) => item.kind === 'column')) {
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

  const plan: QueryPlan = { select }
  if (where !== undefined) {
    plan.where = where
  }
  if (groupBy !== undefined) {
    plan.groupBy = groupBy
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

function hasAggregateSelect(plan: QueryPlan): boolean {
  return plan.select.some((item) => item.kind === 'aggregate')
}

function isAggregatePlan(plan: QueryPlan): boolean {
  return plan.groupBy !== undefined || hasAggregateSelect(plan)
}

function projectPlainRows(rows: readonly QueryRow[], select: readonly QuerySelectItem[]): CellValue[] | undefined {
  const values: CellValue[] = []
  for (const row of rows) {
    for (const item of select) {
      if (item.kind !== 'column') {
        return undefined
      }
      values.push(row.values[item.col] ?? emptyValue)
    }
  }
  return values
}

function valueKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Boolean:
      return `b:${value.value ? '1' : '0'}`
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Error:
      return `err:${value.code}`
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? '-0' : String(value.value)}`
    case ValueTag.String:
      return `s:${value.value}`
  }
}

function aggregateHeaderLabel(deps: LookupQueryBuiltinDeps, header: QueryRow | undefined, item: QueryAggregateSelectItem): CellValue {
  const sourceLabel = header?.values[item.col]
  const label = sourceLabel === undefined || sourceLabel.tag === ValueTag.Empty ? `Col${item.col + 1}` : deps.toStringValue(sourceLabel)
  return textValue(`${item.fn} ${label}`)
}

function buildAggregateHeader(
  deps: LookupQueryBuiltinDeps,
  header: QueryRow | undefined,
  select: readonly QuerySelectItem[],
): QueryAggregateOutputRow[] {
  if (header === undefined) {
    return []
  }
  const values = select.map((item) => {
    if (item.kind === 'column') {
      return header.values[item.col] ?? emptyValue
    }
    return aggregateHeaderLabel(deps, header, item)
  })
  return [{ sourceRow: header.sourceRow, values, groupRows: [] }]
}

function evaluateAggregate(deps: LookupQueryBuiltinDeps, rows: readonly QueryRow[], item: QueryAggregateSelectItem): CellValue {
  if (item.fn === 'count') {
    let count = 0
    for (const row of rows) {
      const value = row.values[item.col] ?? emptyValue
      if (value.tag === ValueTag.Error) {
        return value
      }
      if (value.tag !== ValueTag.Empty) {
        count += 1
      }
    }
    return { tag: ValueTag.Number, value: count }
  }

  let total = 0
  for (const row of rows) {
    const value = row.values[item.col] ?? emptyValue
    if (value.tag === ValueTag.Error) {
      return value
    }
    if (value.tag === ValueTag.Empty) {
      continue
    }
    if (value.tag !== ValueTag.Number) {
      return deps.errorValue(ErrorCode.Value)
    }
    total += value.value
  }
  return { tag: ValueTag.Number, value: total }
}

function evaluateAggregateItem(deps: LookupQueryBuiltinDeps, rows: readonly QueryRow[], item: QuerySelectItem): CellValue {
  if (item.kind === 'column') {
    return rows[0]?.values[item.col] ?? emptyValue
  }
  return evaluateAggregate(deps, rows, item)
}

function groupBodyRows(body: readonly QueryRow[], groupBy: readonly number[] | undefined): QueryRow[][] {
  if (groupBy === undefined) {
    return body.length === 0 ? [[]] : [Array.from(body)]
  }

  const groups = new Map<string, QueryRow[]>()
  for (const row of body) {
    const key = groupBy.map((col) => valueKey(row.values[col] ?? emptyValue)).join('\u001f')
    const existing = groups.get(key)
    if (existing) {
      existing.push(row)
    } else {
      groups.set(key, [row])
    }
  }
  return Array.from(groups.values())
}

function buildAggregateBody(
  deps: LookupQueryBuiltinDeps,
  body: readonly QueryRow[],
  plan: QueryPlan,
): QueryAggregateOutputRow[] | CellValue {
  const outputRows: QueryAggregateOutputRow[] = []
  for (const groupRows of groupBodyRows(body, plan.groupBy)) {
    const values: CellValue[] = []
    for (const item of plan.select) {
      const value = evaluateAggregateItem(deps, groupRows, item)
      if (value.tag === ValueTag.Error) {
        return value
      }
      values.push(value)
    }
    outputRows.push({ sourceRow: groupRows[0]?.sourceRow ?? 0, values, groupRows })
  }
  return outputRows
}

function sortPlainRows(deps: LookupQueryBuiltinDeps, body: QueryRow[], order: QueryOrder): QueryRow[] | CellValue {
  if (order.item.kind !== 'column') {
    return deps.errorValue(ErrorCode.Value)
  }

  let sortError: CellValue | undefined
  const sorted = body.toSorted((left, right) => {
    const leftValue = left.values[order.item.col] ?? emptyValue
    const rightValue = right.values[order.item.col] ?? emptyValue
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
  return sortError ?? sorted
}

function sortAggregateRows(
  deps: LookupQueryBuiltinDeps,
  body: QueryAggregateOutputRow[],
  order: QueryOrder,
): QueryAggregateOutputRow[] | CellValue {
  let sortError: CellValue | undefined
  const sorted = body.toSorted((left, right) => {
    const leftValue = evaluateAggregateItem(deps, left.groupRows, order.item)
    const rightValue = evaluateAggregateItem(deps, right.groupRows, order.item)
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
  return sortError ?? sorted
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

  if (isAggregatePlan(plan)) {
    const aggregateHeader = buildAggregateHeader(deps, header[0], plan.select)
    let aggregateBody = buildAggregateBody(deps, body, plan)
    if (!Array.isArray(aggregateBody)) {
      return aggregateBody
    }
    if (plan.orderBy) {
      aggregateBody = sortAggregateRows(deps, aggregateBody, plan.orderBy)
      if (!Array.isArray(aggregateBody)) {
        return aggregateBody
      }
    }
    const offset = plan.offset ?? 0
    const limitedBody = plan.limit === undefined ? aggregateBody.slice(offset) : aggregateBody.slice(offset, offset + plan.limit)
    const outputRows = [...aggregateHeader, ...limitedBody]
    return deps.arrayResult(
      outputRows.flatMap((row) => row.values),
      outputRows.length,
      plan.select.length,
    )
  }

  if (plan.orderBy) {
    const sorted = sortPlainRows(deps, body, plan.orderBy)
    if (!Array.isArray(sorted)) {
      return sorted
    }
    body = sorted
  }

  const offset = plan.offset ?? 0
  const limitedBody = plan.limit === undefined ? body.slice(offset) : body.slice(offset, offset + plan.limit)
  const outputRows = [...header, ...limitedBody]
  const values = projectPlainRows(outputRows, plan.select)
  if (values === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  return deps.arrayResult(values, outputRows.length, plan.select.length)
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
