import { uiSameCorpusWorkloadMutatesWorkbook, type UiResponsivenessSameCorpusWorkload } from './ui-responsiveness-same-corpus-workloads.ts'

export async function withSameCorpusMutationFailureRestore<T>(args: {
  readonly workload: UiResponsivenessSameCorpusWorkload
  readonly run: () => Promise<T>
  readonly restore: () => Promise<void>
  readonly reselectTarget?: (() => Promise<void>) | undefined
}): Promise<T> {
  try {
    return await args.run()
  } catch (error: unknown) {
    if (uiSameCorpusWorkloadMutatesWorkbook(args.workload)) {
      await args.restore().catch(() => undefined)
      await args.reselectTarget?.().catch(() => undefined)
    }
    throw error
  }
}
