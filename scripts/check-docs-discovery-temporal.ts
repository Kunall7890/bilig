import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

export interface TemporalDiscoveryCheckContext {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly llmsFull: string
  readonly scopedWorkpaperPackageReadme: string
}

export async function requireTemporalWorkpaperActivityDiscovery(context: TemporalDiscoveryCheckContext): Promise<void> {
  const temporalWorkpaperActivity = await readFile(join(context.docsRoot, 'temporal-workpaper-activity.md'), 'utf8')

  await Promise.all(
    [
      'README.md',
      'package.json',
      'tsconfig.json',
      'src/activities.ts',
      'src/workflows.ts',
      'src/smoke.ts',
      'scripts/check-temporal-boundary.ts',
    ].map((sourceFile) => requireFile(join(context.repoRoot, 'examples', 'temporal-workpaper-activity', sourceFile))),
  )

  requireIncludes(temporalWorkpaperActivity, 'examples/temporal-workpaper-activity', 'docs/temporal-workpaper-activity.md')
  requireIncludes(temporalWorkpaperActivity, 'proxyActivities<TemporalWorkPaperActivities>', 'docs/temporal-workpaper-activity.md')
  requireIncludes(temporalWorkpaperActivity, 'WorkflowReplayer', 'docs/temporal-workpaper-activity.md')
  requireIncludes(
    temporalWorkpaperActivity,
    'Temporal owns durable orchestration, retries, workflow history',
    'docs/temporal-workpaper-activity.md',
  )
  requireIncludes(context.llmsFull, 'Temporal WorkPaper Activity', 'docs/llms-full.txt')
  requireIncludes(context.scopedWorkpaperPackageReadme, 'Temporal TypeScript Activity decisions', 'packages/workpaper/README.md')
  requireIncludes(
    context.scopedWorkpaperPackageReadme,
    'https://proompteng.github.io/bilig/temporal-workpaper-activity.html',
    'packages/workpaper/README.md',
  )
}
