import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

const scriptPath = new URL('../src/workpaper-script.ts', import.meta.url)
const outDir = new URL('../.tmp/', import.meta.url)
const outPath = new URL('windmill-hub-script-url.txt', outDir)

const content = await readFile(scriptPath, 'utf8')
const schema = {
  type: 'object',
  properties: {
    quantity: {
      type: 'number',
      description: 'New quote quantity to write into Inputs!B2.',
      default: 18,
    },
    unitPrice: {
      type: 'number',
      description: 'Unit price used by the quote formulas.',
      default: 125,
    },
    discountRate: {
      type: 'number',
      description: 'Discount rate as a decimal, for example 0.1 for 10%.',
      default: 0.1,
    },
    taxRate: {
      type: 'number',
      description: 'Tax rate as a decimal, for example 0.08 for 8%.',
      default: 0.08,
    },
    unitCost: {
      type: 'number',
      description: 'Unit cost used by the margin formula.',
      default: 52,
    },
    previousQuantity: {
      type: 'number',
      description: 'Quantity used to build the before-readback state.',
      default: 12,
    },
  },
}

const state = {
  content,
  summary: 'Calculate quote fields with Bilig WorkPaper formulas',
  description: [
    'Windmill script for quote, payout, or import-check workflows that keep formulas reviewable as workbook cells.',
    '',
    'The script builds a Bilig WorkPaper, edits Inputs!B2, reads dependent formula values, exports WorkPaper JSON, restores it, and returns both a patch object and proof. Keep the proof in logs or an approval step when the next workflow action needs calculated readback evidence.',
    '',
    'Run locally from the Bilig repo with: cd examples/windmill-workpaper-script && pnpm install --ignore-workspace --lockfile=false && pnpm run smoke',
  ].join('\n'),
  kind: 'script',
  language: 'bun',
  schema,
  lock: '',
}

const encoded = Buffer.from(encodeURIComponent(JSON.stringify(state))).toString('base64')
const url = `https://hub.windmill.dev/scripts/add#${encoded}`

await mkdir(outDir, { recursive: true })
await writeFile(outPath, `${url}\n`)

console.log(url)
console.error(`Wrote ${join('.tmp', 'windmill-hub-script-url.txt')}`)
