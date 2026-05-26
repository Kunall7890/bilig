import { readFile } from 'node:fs/promises'

const dag = await readFile(new URL('../dags/bilig_workpaper_quote_dag.py', import.meta.url), 'utf8')
const script = await readFile(new URL('../workpaper-quote.ts', import.meta.url), 'utf8')

for (const needle of [
  'from airflow.sdk import dag, task',
  'from airflow.decorators import dag, task',
  '@dag(',
  'dag_id="bilig_workpaper_quote"',
  '@task(retries=2)',
  'verify_formula_proof',
  '"npx"',
  '"--no-install"',
  '"tsx"',
  'workpaper-quote.ts',
  'workpaper-proof.json',
]) {
  if (!dag.includes(needle)) {
    throw new Error(`bilig_workpaper_quote_dag.py is missing ${needle}`)
  }
}

for (const needle of ['WorkPaper.buildFromSheets', 'exportWorkPaperDocument', 'afterRestore', 'verified']) {
  if (!script.includes(needle)) {
    throw new Error(`workpaper-quote.ts is missing ${needle}`)
  }
}

console.log('Airflow DAG proof files are wired.')
