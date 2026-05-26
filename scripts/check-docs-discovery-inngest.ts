import { join } from 'node:path'

import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

export async function requireInngestWorkpaperStepExampleFiles(repoRoot: string): Promise<void> {
  await Promise.all(
    [
      'README.md',
      'package.json',
      'tsconfig.json',
      'src/workpaper-quote.ts',
      'src/inngest-workpaper-function.ts',
      'src/smoke.ts',
      'scripts/check-inngest-recipe.ts',
    ].map((sourceFile) => requireFile(join(repoRoot, 'examples', 'inngest-workpaper-step', sourceFile))),
  )
}

export function requireInngestWorkpaperStepDiscovery(args: {
  readonly inngestWorkpaperStep: string
  readonly llms: string
  readonly llmsFull: string
  readonly scopedWorkpaperPackageReadme: string
}): void {
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/inngest-workpaper-step.html', 'docs/llms.txt')
  requireIncludes(args.inngestWorkpaperStep, 'examples/inngest-workpaper-step', 'docs/inngest-workpaper-step.md')
  requireIncludes(args.inngestWorkpaperStep, "step.run('calculate-workpaper-quote'", 'docs/inngest-workpaper-step.md')
  requireIncludes(args.inngestWorkpaperStep, 'Inngest owns event delivery, durable step execution', 'docs/inngest-workpaper-step.md')
  requireIncludes(args.llmsFull, 'Inngest WorkPaper Step', 'docs/llms-full.txt')
  requireIncludes(args.scopedWorkpaperPackageReadme, 'Inngest durable step fields', 'packages/workpaper/README.md')
}
