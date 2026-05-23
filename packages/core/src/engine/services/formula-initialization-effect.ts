import { Effect } from 'effect'
import { EngineMutationError } from '../errors.js'
import { mutationErrorMessage } from './formula-initialization-predicates.js'

export function formulaInitializationMutationEffect(message: string, run: () => void) {
  return Effect.try({
    try: run,
    catch: (cause) =>
      new EngineMutationError({
        message: mutationErrorMessage(message, cause),
        cause,
      }),
  })
}
