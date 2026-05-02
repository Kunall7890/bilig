import { useWorkbookGridInteractionRuntime } from './useWorkbookGridInteractionRuntime.js'
import { useWorkbookGridPaneRenderRuntime } from './useWorkbookGridPaneRenderRuntime.js'

type InteractionRuntimeInput = Parameters<typeof useWorkbookGridInteractionRuntime>[0]
type DrawRuntimeInput = Parameters<typeof useWorkbookGridPaneRenderRuntime>[0]
type InteractionRuntimeResult = ReturnType<typeof useWorkbookGridInteractionRuntime>
type DrawRuntimeResult = ReturnType<typeof useWorkbookGridPaneRenderRuntime>
type RenderPipelineRuntimeInput = InteractionRuntimeInput & Omit<DrawRuntimeInput, 'requiresLiveViewportState'>

export function useWorkbookGridRenderPipelineRuntime(input: RenderPipelineRuntimeInput): InteractionRuntimeResult & DrawRuntimeResult {
  const interactionState = useWorkbookGridInteractionRuntime(input)
  const paneState = useWorkbookGridPaneRenderRuntime({
    ...input,
    requiresLiveViewportState: interactionState.requiresLiveViewportState,
  })

  return {
    ...interactionState,
    ...paneState,
  }
}
