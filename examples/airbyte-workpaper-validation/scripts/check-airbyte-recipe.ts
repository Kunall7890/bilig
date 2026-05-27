import { readFile } from 'node:fs/promises'

const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8')
const fixture = await readFile(new URL('../fixtures/orders-airbyte-messages.jsonl', import.meta.url), 'utf8')
const helperSource = await readFile(new URL('../src/airbyte-workpaper-validation.ts', import.meta.url), 'utf8')
const smokeSource = await readFile(new URL('../src/smoke.ts', import.meta.url), 'utf8')

for (const needle of [
  'Airbyte WorkPaper Validation',
  'fixtures/orders-airbyte-messages.jsonl',
  'https://docs.airbyte.com/platform/understanding-airbyte/airbyte-protocol',
  'https://docs.airbyte.com/platform/using-airbyte/core-concepts/sync-modes/incremental-append-deduped',
  'Airbyte owns the sync and checkpoint semantics.',
]) {
  if (!readme.includes(needle)) {
    throw new Error(`README.md is missing ${needle}`)
  }
}

for (const needle of ['"type":"RECORD"', '"type":"STATE"', '"stream":"orders"', '"cursor":"2026-05-27T10:10:00Z"']) {
  if (!fixture.includes(needle)) {
    throw new Error(`fixtures/orders-airbyte-messages.jsonl is missing ${needle}`)
  }
}

for (const needle of [
  '@bilig/workpaper',
  'readAirbyteMessagesFromJsonl',
  'validateAirbyteOrdersWithWorkPaper',
  'WorkPaper.buildFromSheets',
  'exportWorkPaperDocument',
  'afterRestore',
  'persistedDocumentBytes',
  'validation_passed',
]) {
  if (!helperSource.includes(needle)) {
    throw new Error(`src/airbyte-workpaper-validation.ts is missing ${needle}`)
  }
}

for (const needle of [
  'orders-airbyte-messages.jsonl',
  'expectedPaidAmount: 301.75',
  'expectedRecordCount: 4',
  'stateCursorMatchesRecords',
  'verified',
]) {
  if (!smokeSource.includes(needle)) {
    throw new Error(`src/smoke.ts is missing ${needle}`)
  }
}

console.log('Airbyte WorkPaper validation recipe is wired.')
