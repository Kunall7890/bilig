export function rethrowFatalFormulaBindingError(error: unknown): void {
  if (isFormulaBindingTimeoutError(error)) {
    throw error
  }
}

function isFormulaBindingTimeoutError(error: unknown): boolean {
  let current: unknown = error
  let depth = 0
  while (typeof current === 'object' && current !== null && depth < 16) {
    if (current instanceof Error && current.name === 'EngineEvaluationTimeoutError') {
      return true
    }
    current = Reflect.get(current, 'cause')
    depth += 1
  }
  return false
}
