import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const previewWidth = 1280
const previewHeight = 640
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const backgroundPath = join(assetsRoot, 'bilig-social-background.png')
const outputPath = join(assetsRoot, 'github-social-preview.png')
const checkMode = process.argv.includes('--check')

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 16 * 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error !== null) {
          const message = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : String(stderr)
          reject(new Error(`${file} failed: ${message.trim() || error.message}`))
          return
        }
        resolve(Buffer.from(stdout))
      },
    )
  })
}

async function buildSvg(): Promise<string> {
  const background = await readFile(backgroundPath)
  const backgroundUri = `data:image/png;base64,${background.toString('base64')}`

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${previewWidth}" height="${previewHeight}" viewBox="0 0 ${previewWidth} ${previewHeight}">
  <image href="${backgroundUri}" x="0" y="0" width="${previewWidth}" height="${previewHeight}" preserveAspectRatio="xMidYMid slice"/>
</svg>`
}

async function renderPreview(): Promise<Buffer> {
  const svg = await buildSvg()
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-social-preview-'))
  const svgPath = join(tempRoot, 'preview.svg')
  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', [
      '--format=png',
      '--width',
      String(previewWidth),
      '--height',
      String(previewHeight),
      svgPath,
    ])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

const image = await renderPreview()

if (checkMode) {
  const existing = await readFile(outputPath)
  if (!existing.equals(image)) {
    throw new Error(`${outputPath} is stale. Run pnpm docs:social-preview:generate.`)
  }
  console.log(`social preview is current: ${outputPath}`)
} else {
  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(outputPath, image)
  console.log(`wrote ${outputPath}`)
}
