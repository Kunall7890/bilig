import { Cause, Effect, Exit } from 'effect'

export function runEngineEffect<Success, Failure>(effect: Effect.Effect<Success, Failure>): Success {
  const exit = Effect.runSyncExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}

export async function runEngineEffectPromise<Success, Failure>(effect: Effect.Effect<Success, Failure>): Promise<Success> {
  const exit = await Effect.runPromiseExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}
