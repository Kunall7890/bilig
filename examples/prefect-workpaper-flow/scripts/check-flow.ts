import { readFile } from 'node:fs/promises'

const flow = await readFile(new URL('../flow.py', import.meta.url), 'utf8')
const script = await readFile(new URL('../workpaper-quote.ts', import.meta.url), 'utf8')

for (const needle of [
  'from prefect import flow, task',
  '@flow(name="bilig-workpaper-quote")',
  '@task(retries=2, retry_delay_seconds=5)',
  '"npx"',
  '"tsx"',
  'workpaper-quote.ts',
  'workpaper-proof.json',
]) {
  if (!flow.includes(needle)) {
    throw new Error(`flow.py is missing ${needle}`)
  }
}

for (const needle of ['WorkPaper.buildFromSheets', 'exportWorkPaperDocument', 'afterRestore', 'verified']) {
  if (!script.includes(needle)) {
    throw new Error(`workpaper-quote.ts is missing ${needle}`)
  }
}

console.log('Prefect flow proof files are wired.')
