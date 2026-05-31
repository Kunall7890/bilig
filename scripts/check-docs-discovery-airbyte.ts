import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

const repoRoot = join(import.meta.dirname, '..')
const docsRoot = join(repoRoot, 'docs')
const exampleRoot = join(repoRoot, 'examples', 'airbyte-workpaper-validation')

const [readme, llms, llmsFull, workpaperReadme, airbyteDoc, fixture, globalFixture, helperSource, smokeSource] = await Promise.all([
  readFile(join(repoRoot, 'README.md'), 'utf8'),
  readFile(join(docsRoot, 'llms.txt'), 'utf8'),
  readFile(join(docsRoot, 'llms-full.txt'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'workpaper', 'README.md'), 'utf8'),
  readFile(join(docsRoot, 'airbyte-workpaper-validation.md'), 'utf8'),
  readFile(join(exampleRoot, 'fixtures', 'orders-airbyte-messages.jsonl'), 'utf8'),
  readFile(join(exampleRoot, 'fixtures', 'orders-airbyte-global-state-messages.jsonl'), 'utf8'),
  readFile(join(exampleRoot, 'src', 'airbyte-workpaper-validation.ts'), 'utf8'),
  readFile(join(exampleRoot, 'src', 'smoke.ts'), 'utf8'),
])

await Promise.all(
  [
    'README.md',
    'package.json',
    'tsconfig.json',
    'fixtures/orders-airbyte-messages.jsonl',
    'fixtures/orders-airbyte-global-state-messages.jsonl',
    'src/airbyte-workpaper-validation.ts',
    'src/smoke.ts',
    'scripts/check-airbyte-recipe.ts',
  ].map((sourceFile) => requireFile(join(exampleRoot, sourceFile))),
)

for (const source of [readme, llms, llmsFull, workpaperReadme]) {
  requireIncludes(source, 'airbyte-workpaper-validation', 'Airbyte discovery surfaces')
}

for (const needle of [
  'Airbyte WorkPaper Validation',
  'examples/airbyte-workpaper-validation',
  'fixtures/orders-airbyte-messages.jsonl',
  'fixtures/orders-airbyte-global-state-messages.jsonl',
  'GLOBAL',
  'Airbyte owns extraction, replication, sync mode selection',
  'https://docs.airbyte.com/platform/understanding-airbyte/airbyte-protocol',
  'https://docs.airbyte.com/platform/using-airbyte/core-concepts/sync-modes/incremental-append-deduped',
]) {
  requireIncludes(airbyteDoc, needle, 'docs/airbyte-workpaper-validation.md')
}

for (const needle of ['"type":"RECORD"', '"type":"STATE"', '"stream":"orders"', '"cursor":"2026-05-27T10:10:00Z"']) {
  requireIncludes(fixture, needle, 'examples/airbyte-workpaper-validation/fixtures/orders-airbyte-messages.jsonl')
}

for (const needle of ['"state_type":"GLOBAL"', '"global"', '"stream_states"', '"stream_descriptor"', '"stream_state"']) {
  requireIncludes(globalFixture, needle, 'examples/airbyte-workpaper-validation/fixtures/orders-airbyte-global-state-messages.jsonl')
}

for (const needle of [
  'readAirbyteMessagesFromJsonl',
  'validateAirbyteOrdersWithWorkPaper',
  'readGlobalStateCursor',
  'WorkPaper.buildFromSheets',
  'exportWorkPaperDocument',
  'state_type',
  'afterRestore',
  'persistedDocumentBytes',
  'validation_passed',
]) {
  requireIncludes(helperSource, needle, 'examples/airbyte-workpaper-validation/src/airbyte-workpaper-validation.ts')
}

for (const needle of [
  'orders-airbyte-global-state-messages.jsonl',
  'expectedPaidAmount: 301.75',
  'expectedRecordCount: 4',
  'stateCursorMatchesRecords',
  'verified',
]) {
  requireIncludes(smokeSource, needle, 'examples/airbyte-workpaper-validation/src/smoke.ts')
}

console.log('Airbyte WorkPaper validation discovery is wired.')
