import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const rootActionPath = join(repoRoot, 'action.yml')
const nestedActionPath = join(repoRoot, 'actions', 'xlsx-cache-doctor', 'action.yml')

const [rootAction, nestedAction] = await Promise.all([readFile(rootActionPath, 'utf8'), readFile(nestedActionPath, 'utf8')])

if (rootAction !== nestedAction) {
  throw new Error(
    [
      'Root action.yml must stay byte-for-byte aligned with actions/xlsx-cache-doctor/action.yml.',
      'Edit both files together so Marketplace and subdirectory action users get the same inputs, outputs, and behavior.',
    ].join('\n'),
  )
}
