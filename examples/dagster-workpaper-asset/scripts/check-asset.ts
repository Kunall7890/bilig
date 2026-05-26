import { readFile } from 'node:fs/promises'

const asset = await readFile(new URL('../defs/bilig_workpaper_asset.py', import.meta.url), 'utf8')
const script = await readFile(new URL('../workpaper-asset.ts', import.meta.url), 'utf8')

for (const needle of [
  'import dagster as dg',
  '@dg.asset(compute_kind="javascript")',
  'dg.PipesSubprocessClient',
  'bilig_workpaper_quote_asset',
  '"npx"',
  '"--no-install"',
  '"tsx"',
  'workpaper-asset.ts',
  'workpaper-proof.json',
]) {
  if (!asset.includes(needle)) {
    throw new Error(`bilig_workpaper_asset.py is missing ${needle}`)
  }
}

for (const needle of [
  'WorkPaper.buildFromSheets',
  'DAGSTER_PIPES_CONTEXT',
  'DAGSTER_PIPES_MESSAGES',
  'report_asset_materialization',
  'exportWorkPaperDocument',
  'afterRestore',
  'verified',
]) {
  if (!script.includes(needle)) {
    throw new Error(`workpaper-asset.ts is missing ${needle}`)
  }
}

console.log('Dagster WorkPaper asset proof files are wired.')
