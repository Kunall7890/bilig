import { readFile } from 'node:fs/promises'

const workflow = await readFile(new URL('../src/workflows.ts', import.meta.url), 'utf8')
const activity = await readFile(new URL('../src/activities.ts', import.meta.url), 'utf8')
const smoke = await readFile(new URL('../src/smoke.ts', import.meta.url), 'utf8')
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8')

for (const forbidden of ['@bilig/workpaper', 'WorkPaper.buildFromSheets', 'writeFile', 'readFile']) {
  if (workflow.includes(forbidden)) {
    throw new Error(`src/workflows.ts must not include ${forbidden}`)
  }
}

for (const needle of [
  '@temporalio/workflow',
  'proxyActivities<TemporalWorkPaperActivities>',
  'startToCloseTimeout',
  'maximumAttempts: 3',
  'quoteApprovalWorkflow',
  'calculateWorkPaperQuoteActivity',
]) {
  if (!workflow.includes(needle)) {
    throw new Error(`src/workflows.ts is missing ${needle}`)
  }
}

for (const needle of [
  '@bilig/workpaper',
  'WorkPaper.buildFromSheets',
  'exportWorkPaperDocument',
  'afterRestore',
  'persistedDocumentBytes',
  'workflowImportsWorkPaper: false',
  'activityOwnsWorkPaper: true',
]) {
  if (!activity.includes(needle)) {
    throw new Error(`src/activities.ts is missing ${needle}`)
  }
}

for (const needle of ['MockActivityEnvironment', 'isExpectedProof', "Reflect.get(patch, 'total') === 2187"]) {
  if (!smoke.includes(needle)) {
    throw new Error(`src/smoke.ts is missing ${needle}`)
  }
}

for (const needle of [
  'Temporal WorkPaper Activity',
  'Workflow code imports only `@temporalio/workflow`',
  'cd examples/temporal-workpaper-activity',
  'pnpm run smoke',
  'WorkflowReplayer',
]) {
  if (!readme.includes(needle)) {
    throw new Error(`README.md is missing ${needle}`)
  }
}

console.log('Temporal WorkPaper Activity boundary is wired.')
