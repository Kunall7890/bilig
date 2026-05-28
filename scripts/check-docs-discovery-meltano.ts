import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

const repoRoot = join(import.meta.dirname, '..')
const docsRoot = join(repoRoot, 'docs')
const exampleRoot = join(repoRoot, 'examples', 'meltano-workpaper-utility')

const [readme, llms, llmsFull, index, workpaperReadme, meltanoDoc, exampleReadme, meltanoYml, hubDefinition, fixture, source] =
  await Promise.all([
    readFile(join(repoRoot, 'README.md'), 'utf8'),
    readFile(join(docsRoot, 'llms.txt'), 'utf8'),
    readFile(join(docsRoot, 'llms-full.txt'), 'utf8'),
    readFile(join(docsRoot, 'index.html'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'workpaper', 'README.md'), 'utf8'),
    readFile(join(docsRoot, 'meltano-workpaper-utility.md'), 'utf8'),
    readFile(join(exampleRoot, 'README.md'), 'utf8'),
    readFile(join(exampleRoot, 'meltano.yml'), 'utf8'),
    readFile(join(exampleRoot, 'meltano-hub-utility-definition.yml'), 'utf8'),
    readFile(join(exampleRoot, 'fixtures', 'orders.jsonl'), 'utf8'),
    readFile(join(exampleRoot, 'meltano-workpaper-validator.ts'), 'utf8'),
  ])

await Promise.all(
  [
    'README.md',
    'package.json',
    'tsconfig.json',
    'meltano.yml',
    'meltano-hub-utility-definition.yml',
    'fixtures/orders.jsonl',
    'meltano-workpaper-validator.ts',
    'scripts/check-meltano-recipe.ts',
  ].map((sourceFile) => requireFile(join(exampleRoot, sourceFile))),
)

for (const surface of [readme, llms, llmsFull, index, workpaperReadme]) {
  requireIncludes(surface, 'meltano-workpaper-utility', 'Meltano discovery surfaces')
}

for (const needle of [
  'Meltano WorkPaper Utility',
  'examples/meltano-workpaper-utility',
  'meltano invoke bilig-workpaper-validator:validate',
  'meltano-hub-utility-definition.yml',
  'https://docs.meltano.com/concepts/plugins/#custom-utilities',
  'https://docs.meltano.com/reference/command-line-interface/#invoke',
  'https://docs.meltano.com/reference/plugin-definition-syntax/#commands',
]) {
  requireIncludes(meltanoDoc, needle, 'docs/meltano-workpaper-utility.md')
  requireIncludes(exampleReadme, needle, 'examples/meltano-workpaper-utility/README.md')
}

for (const needle of ['utilities:', 'executable: npx', 'commands:', 'validate:', 'bilig-workpaper-validator:validate']) {
  requireIncludes(meltanoYml, needle, 'examples/meltano-workpaper-utility/meltano.yml')
}

for (const needle of ['variant: proompteng', 'maintenance_status: active', 'usage:', 'next_steps:']) {
  requireIncludes(hubDefinition, needle, 'examples/meltano-workpaper-utility/meltano-hub-utility-definition.yml')
}

for (const needle of ['"status":"paid"', '"status":"rejected"', '"amount"']) {
  requireIncludes(fixture, needle, 'examples/meltano-workpaper-utility/fixtures/orders.jsonl')
}

for (const needle of [
  '@bilig/workpaper',
  'WorkPaper.buildFromSheets',
  'exportWorkPaperDocument',
  'afterRestore',
  'persistedDocumentBytes',
  'validation_passed',
]) {
  requireIncludes(source, needle, 'examples/meltano-workpaper-utility/meltano-workpaper-validator.ts')
}

console.log('Meltano WorkPaper utility discovery is wired.')
