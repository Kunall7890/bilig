import { readFile } from 'node:fs/promises'

const flow = await readFile(new URL('../flow.yml', import.meta.url), 'utf8')
const blueprint = await readFile(new URL('../blueprint.yaml', import.meta.url), 'utf8')
const script = await readFile(new URL('../kestra-workpaper-flow.ts', import.meta.url), 'utf8')

for (const needle of [
  'io.kestra.plugin.scripts.node.Commands',
  'namespaceFiles:',
  'containerImage: node:24-slim',
  'npm install @bilig/workpaper@latest',
  'npx tsx kestra-workpaper-flow.ts',
  'outputFiles:',
  'workpaper-proof.json',
]) {
  if (!flow.includes(needle)) {
    throw new Error(`flow.yml is missing ${needle}`)
  }
}

for (const needle of [
  'id: bilig-workpaper-formula-readback',
  'extend:',
  'title: Verify spreadsheet-style quote formulas with Bilig WorkPaper',
  'io.kestra.plugin.scripts.node.Script',
  'containerImage: node:24-slim',
  '"@bilig/workpaper": "latest"',
  'outputFiles:',
  'workpaper-proof.json',
  'verified',
  'metaDescription: Run Bilig WorkPaper formulas in a Kestra Node task',
]) {
  if (!blueprint.includes(needle)) {
    throw new Error(`blueprint.yaml is missing ${needle}`)
  }
}

for (const needle of ['WorkPaper.buildFromSheets', 'exportWorkPaperDocument', 'afterRestore', 'verified']) {
  if (!script.includes(needle)) {
    throw new Error(`kestra-workpaper-flow.ts is missing ${needle}`)
  }
}

console.log('Kestra flow proof files are wired.')
