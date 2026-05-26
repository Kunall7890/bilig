import { readFile } from 'node:fs/promises'

const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8')
const functionSource = await readFile(new URL('../src/inngest-workpaper-function.ts', import.meta.url), 'utf8')
const helperSource = await readFile(new URL('../src/workpaper-quote.ts', import.meta.url), 'utf8')
const smokeSource = await readFile(new URL('../src/smoke.ts', import.meta.url), 'utf8')

for (const needle of [
  'Inngest WorkPaper Step',
  'step.run()',
  'pnpm run smoke',
  'https://www.inngest.com/docs/reference/typescript/v3/functions/step-run',
]) {
  if (!readme.includes(needle)) {
    throw new Error(`README.md is missing ${needle}`)
  }
}

for (const needle of [
  "import { Inngest } from 'inngest'",
  'new Inngest',
  'createFunction',
  "triggers: [{ event: 'bilig/quote.requested' }]",
  "step.run('calculate-workpaper-quote'",
  'result.proof.verified',
]) {
  if (!functionSource.includes(needle)) {
    throw new Error(`src/inngest-workpaper-function.ts is missing ${needle}`)
  }
}

for (const needle of [
  '@bilig/workpaper',
  'WorkPaper.buildFromSheets',
  'exportWorkPaperDocument',
  'afterRestore',
  'persistedDocumentBytes',
]) {
  if (!helperSource.includes(needle)) {
    throw new Error(`src/workpaper-quote.ts is missing ${needle}`)
  }
}

for (const needle of ['calculateWorkPaperQuote', "Reflect.get(patch, 'total') === 2187", 'verified']) {
  if (!smokeSource.includes(needle)) {
    throw new Error(`src/smoke.ts is missing ${needle}`)
  }
}

console.log('Inngest WorkPaper step recipe is wired.')
