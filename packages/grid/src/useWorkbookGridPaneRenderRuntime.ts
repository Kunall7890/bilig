import { useWorkbookGridDrawRuntime } from './useWorkbookGridDrawRuntime.js'

type UseWorkbookGridPaneRenderRuntimeInput = Parameters<typeof useWorkbookGridDrawRuntime>[0]
type UseWorkbookGridPaneRenderRuntimeResult = ReturnType<typeof useWorkbookGridDrawRuntime>

export function useWorkbookGridPaneRenderRuntime(input: UseWorkbookGridPaneRenderRuntimeInput): UseWorkbookGridPaneRenderRuntimeResult {
  return useWorkbookGridDrawRuntime(input)
}
