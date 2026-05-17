import { stringifyWorkbookAgentUiContextSemanticKey, type WorkbookAgentUiContext } from '@bilig/contracts'

export function areWorkbookAgentUiContextsSemanticallyEqual(
  left: WorkbookAgentUiContext | null,
  right: WorkbookAgentUiContext | null,
): boolean {
  return stringifyWorkbookAgentUiContextSemanticKey(left) === stringifyWorkbookAgentUiContextSemanticKey(right)
}
