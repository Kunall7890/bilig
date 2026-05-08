import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'

import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'

const rootDir = fileURLToPath(new URL('..', import.meta.url))

try {
  assertLocalCiResourceGuardAllowsRun(rootDir, process.env, { runLabel: 'pre-push lint' })
  const lintExitCode = await run('pnpm', ['lint'])
  process.exit(lintExitCode)
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
}

function run(command: string, args: readonly string[]): Promise<number> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: rootDir,
      env: process.env,
      stdio: 'inherit',
    })

    child.on('error', reject)
    child.on('close', (code, signal) => {
      if (signal) {
        reject(new Error(`${command} ${args.join(' ')} terminated by ${signal}`))
        return
      }
      resolve(code ?? 1)
    })
  })
}
