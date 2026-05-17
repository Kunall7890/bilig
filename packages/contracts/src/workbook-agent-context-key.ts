import type { WorkbookAgentRenderedRange, WorkbookAgentUiContext } from './index.js'

const WORKBOOK_AGENT_STRING_VALUE_TAG = 3

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeRenderedValueForContextKey(value: unknown): unknown {
  if (!isRecord(value) || value['tag'] !== WORKBOOK_AGENT_STRING_VALUE_TAG) {
    return value
  }
  return {
    tag: value['tag'],
    value: value['value'],
    stringId: 0,
  }
}

function normalizeRenderedRangeForContextKey(range: WorkbookAgentRenderedRange | null): WorkbookAgentRenderedRange | null {
  if (range === null) {
    return null
  }
  return {
    ...range,
    rows: range.rows.map((row) =>
      row.map((cell) => ({
        ...cell,
        value: normalizeRenderedValueForContextKey(cell.value),
      })),
    ),
  }
}

export function stringifyWorkbookAgentUiContextSemanticKey(context: WorkbookAgentUiContext | null): string {
  if (context === null) {
    return 'null'
  }
  const rendered = context.rendered
  return JSON.stringify({
    selection: context.selection,
    viewport: context.viewport,
    rendered:
      rendered === undefined
        ? null
        : {
            selection: normalizeRenderedRangeForContextKey(rendered.selection),
            visibleRange: normalizeRenderedRangeForContextKey(rendered.visibleRange),
          },
  })
}
