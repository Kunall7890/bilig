import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes } from './check-docs-discovery-core.ts'

export async function requireExternalWorkbookRecalcProofDiscovery(args: {
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly xlsxRecalcPackageReadme: string
}): Promise<void> {
  const proof = await readFile(join(args.docsRoot, 'external-workbook-recalc-proof.md'), 'utf8')

  requireIncludes(args.xlsxRecalcPackageReadme, 'external-workbook-recalc-proof.html', 'packages/xlsx-formula-recalc/README.md')
  requireIncludes(args.index, './external-workbook-recalc-proof.html', 'docs/index.html')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/external-workbook-recalc-proof.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/external-workbook-recalc-proof.ts', 'docs/llms.txt')
  for (const required of [
    'title: External workbook recalculation proof in Node.js',
    'canonical_url: https://proompteng.github.io/bilig/external-workbook-recalc-proof.html',
    'npm install @bilig/xlsx-formula-recalc tsx',
    'externalWorkbookMatched',
    'recalculatedExternalLookup',
    'verified',
    'downloadable external-workbook proof script',
  ] as const) {
    requireIncludes(proof, required, 'docs/external-workbook-recalc-proof.md')
  }
}
