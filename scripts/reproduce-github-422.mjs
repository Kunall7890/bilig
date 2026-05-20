import { createHash } from 'node:crypto'
import { performance } from 'node:perf_hooks'

import { WorkPaper } from '../packages/headless/dist/index.js'
import { exportXlsx, importXlsx } from '../packages/headless/dist/xlsx.js'

const [url, filename = 'workbook.xlsx'] = process.argv.slice(2)
if (!url) {
  throw new Error('Usage: node scripts/reproduce-github-422.mjs <xlsx-url> [filename]')
}

const started = performance.now()

function log(stage, extra = {}) {
  console.error(
    JSON.stringify({
      stage,
      elapsedMs: Math.round((performance.now() - started) * 100) / 100,
      rssMb: Math.round((process.memoryUsage().rss / 1024 / 1024) * 100) / 100,
      ...extra,
    }),
  )
}

log('download:start')
const response = await fetch(url)
if (!response.ok) {
  throw new Error(`HTTP ${response.status} while downloading ${url}`)
}
const bytes = new Uint8Array(await response.arrayBuffer())
log('download:done', {
  bytes: bytes.byteLength,
  sha256: createHash('sha256').update(bytes).digest('hex'),
})

globalThis.gc?.()
log('import_xlsx:start')
const importStarted = performance.now()
let imported
try {
  imported = importXlsx(bytes, filename)
} catch (error) {
  log('import_xlsx:failed', {
    importMs: Math.round((performance.now() - importStarted) * 100) / 100,
    errorName: error?.name,
    error: error?.message,
    reason: error?.reason,
    stats: error?.stats,
    limits: error?.limits,
  })
  process.exit(1)
}
log('import_xlsx:done', {
  importMs: Math.round((performance.now() - importStarted) * 100) / 100,
  sheets: imported.snapshot.sheets.length,
  warnings: imported.warnings?.length ?? 0,
})

let cells = 0
let formulaCells = 0
for (const sheet of imported.snapshot.sheets) {
  cells += sheet.cells?.length ?? 0
  for (const cell of sheet.cells ?? []) {
    if (cell.formula !== undefined) {
      formulaCells += 1
    }
  }
}
log('snapshot_scan:done', { cells, formulaCells })

globalThis.gc?.()
log('build_workpaper:start')
const buildStarted = performance.now()
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  useColumnIndex: true,
  evaluationTimeoutMs: 30_000,
})
log('build_workpaper:done', {
  buildMs: Math.round((performance.now() - buildStarted) * 100) / 100,
})

globalThis.gc?.()
log('export_xlsx:start')
const exportStarted = performance.now()
const exported = exportXlsx(workbook.exportSnapshot())
log('export_xlsx:done', {
  bytes: exported.byteLength,
  exportMs: Math.round((performance.now() - exportStarted) * 100) / 100,
})
workbook.dispose?.()
