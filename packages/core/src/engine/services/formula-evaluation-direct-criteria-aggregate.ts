import type { ErrorCode} from '@bilig/protocol';
import { ValueTag, type CellValue } from '@bilig/protocol'
import { directErrorResult } from './formula-evaluation-helpers.js'
import type { RuntimeColumnView } from './runtime-column-store-service.js'

export const firstMatchedAggregateError = (
  view: RuntimeColumnView,
  rows: ArrayLike<number>,
  length: number,
): CellValue | undefined => {
  for (let index = 0; index < length; index += 1) {
    const row = rows[index]!
    if ((view.readTagAt(row) as ValueTag) === ValueTag.Error) {
      return directErrorResult(view.readErrorAt(row) as ErrorCode)
    }
  }
  return undefined
}
