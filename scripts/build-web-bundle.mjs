import { build } from 'vite'
import { fileURLToPath } from 'node:url'

const webRoot = fileURLToPath(new URL('../apps/web/', import.meta.url))
const configFile = fileURLToPath(new URL('../apps/web/vite.config.ts', import.meta.url))

try {
  await build({
    root: webRoot,
    configFile,
  })
  // Vite/Rolldown can leave native worker handles open after a successful
  // production build on macOS. The build promise is the contract here.
  process.exit(0)
} catch (error) {
  console.error(error instanceof Error ? (error.stack ?? error.message) : String(error))
  process.exit(1)
}
