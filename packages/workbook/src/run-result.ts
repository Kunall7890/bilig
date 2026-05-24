import type { WorkbookRunResult } from './result.js'

export function freezeWorkbookRunResult(result: WorkbookRunResult): WorkbookRunResult {
  return freezeData(result, new WeakSet())
}

function freezeData<T>(value: T, seen: WeakSet<object>): T {
  if (typeof value !== 'object' || value === null) {
    return value
  }
  if (seen.has(value)) {
    return value
  }
  seen.add(value)
  Object.values(Object.getOwnPropertyDescriptors(value)).forEach((descriptor) => {
    if ('value' in descriptor) {
      freezeData(descriptor.value, seen)
    }
  })
  return Object.freeze(value)
}
