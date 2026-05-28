import { readFile } from 'node:fs/promises'

const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8')
const meltano = await readFile(new URL('../meltano.yml', import.meta.url), 'utf8')
const hubDefinition = await readFile(new URL('../meltano-hub-utility-definition.yml', import.meta.url), 'utf8')
const fixture = await readFile(new URL('../fixtures/orders.jsonl', import.meta.url), 'utf8')
const script = await readFile(new URL('../meltano-workpaper-validator.ts', import.meta.url), 'utf8')

for (const needle of [
  'Meltano WorkPaper Utility',
  'meltano invoke bilig-workpaper-validator:validate',
  'fixtures/orders.jsonl',
  'meltano-hub-utility-definition.yml',
  'https://docs.meltano.com/concepts/plugins/#custom-utilities',
  'https://docs.meltano.com/reference/command-line-interface/#invoke',
  'https://docs.meltano.com/reference/plugin-definition-syntax/#commands',
]) {
  if (!readme.includes(needle)) {
    throw new Error(`README.md is missing ${needle}`)
  }
}

for (const needle of [
  'version: 1',
  'utilities:',
  'name: bilig-workpaper-validator',
  'executable: npx',
  'commands:',
  'validate:',
  'bilig-workpaper-validator:validate',
]) {
  if (!meltano.includes(needle)) {
    throw new Error(`meltano.yml is missing ${needle}`)
  }
}

for (const needle of [
  'variant: proompteng',
  'maintenance_status: active',
  'commands:',
  'usage:',
  'next_steps:',
  'https://proompteng.github.io/bilig/meltano-workpaper-utility.html',
]) {
  if (!hubDefinition.includes(needle)) {
    throw new Error(`meltano-hub-utility-definition.yml is missing ${needle}`)
  }
}

for (const needle of ['"order_id"', '"amount"', '"status":"paid"', '"status":"rejected"']) {
  if (!fixture.includes(needle)) {
    throw new Error(`fixtures/orders.jsonl is missing ${needle}`)
  }
}

for (const needle of [
  '@bilig/workpaper',
  'WorkPaper.buildFromSheets',
  'meltano invoke bilig-workpaper-validator:validate',
  'exportWorkPaperDocument',
  'afterRestore',
  'persistedDocumentBytes',
  'validation_passed',
]) {
  if (!script.includes(needle)) {
    throw new Error(`meltano-workpaper-validator.ts is missing ${needle}`)
  }
}

console.log('Meltano WorkPaper utility proof files are wired.')
