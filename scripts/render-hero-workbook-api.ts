import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const heroWidth = 1600
const heroHeight = 900
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const outputPath = join(assetsRoot, 'bilig-hero-workbook-api.png')
const svgOutputPath = join(assetsRoot, 'bilig-hero-workbook-api.svg')
const checkMode = process.argv.includes('--check')

function fontFace(family: string, weight: number, data: Buffer): string {
  return String.raw`
    @font-face {
      font-family: '${family}';
      font-weight: ${weight};
      src: url(data:font/woff2;base64,${data.toString('base64')}) format('woff2');
    }`
}

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 24 * 1024 * 1024,
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

function requirePngDimensions(image: Buffer, context: string): void {
  const pngSignature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  if (image.length < 24 || !image.subarray(0, pngSignature.length).equals(pngSignature)) {
    throw new Error(`${context} is not a PNG image`)
  }

  const width = image.readUInt32BE(16)
  const height = image.readUInt32BE(20)
  if (width !== heroWidth || height !== heroHeight) {
    throw new Error(`${context} must be ${heroWidth.toString()}x${heroHeight.toString()}; got ${width.toString()}x${height.toString()}`)
  }
}

async function buildSvg(): Promise<string> {
  const sansRegular = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-400.woff2'))
  const sansMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-500.woff2'))
  const sansSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-600.woff2'))
  const sansBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-700.woff2'))
  const monoMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-500.woff2'))
  const monoSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-600.woff2'))

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${heroWidth}" height="${heroHeight}" viewBox="0 0 ${heroWidth} ${heroHeight}">
  <defs>
    <style>
      ${fontFace('Bilig Sans', 400, sansRegular)}
      ${fontFace('Bilig Sans', 500, sansMedium)}
      ${fontFace('Bilig Sans', 600, sansSemiBold)}
      ${fontFace('Bilig Sans', 700, sansBold)}
      ${fontFace('Bilig Mono', 500, monoMedium)}
      ${fontFace('Bilig Mono', 600, monoSemiBold)}
      text {
        font-family: 'Bilig Sans', Arial, sans-serif;
        letter-spacing: 0;
      }
      .mono {
        font-family: 'Bilig Mono', Menlo, monospace;
      }
    </style>
    <linearGradient id="sheet" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#fffefa"/>
      <stop offset="1" stop-color="#f3f7ef"/>
    </linearGradient>
    <linearGradient id="selectedCell" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#d7f4e2"/>
      <stop offset="1" stop-color="#c9efd8"/>
    </linearGradient>
  </defs>

  <g transform="translate(86 96)">
    <rect x="0" y="0" width="1428" height="66" fill="#fbfbf5" stroke="#cfd9c9"/>
    <rect x="26" y="16" width="104" height="34" rx="4" fill="#edf3e9" stroke="#cfd9c9"/>
    <text x="78" y="39" text-anchor="middle" fill="#596357" font-size="18" font-weight="700">D5</text>
    <path d="M154 0 V66" stroke="#dfe6dc"/>
    <text x="188" y="40" class="mono" fill="#13794b" font-size="22" font-weight="600">=SUM(D2:D4)</text>

    <g transform="translate(0 116)">
      <rect x="0" y="0" width="1428" height="612" fill="url(#sheet)" stroke="#d5ded0"/>
      <rect x="0" y="0" width="1428" height="64" fill="#e8efe5"/>
      <rect x="0" y="0" width="72" height="612" fill="#e8efe5"/>
      <path d="M72 0 V612 M336 0 V612 M600 0 V612 M864 0 V612 M1128 0 V612 M1392 0 V612" stroke="#d5ded0"/>
      <path d="M0 64 H1428 M0 125 H1428 M0 186 H1428 M0 247 H1428 M0 308 H1428 M0 369 H1428 M0 430 H1428 M0 491 H1428 M0 552 H1428" stroke="#d5ded0"/>

      <rect x="336" y="125" width="264" height="61" fill="url(#selectedCell)"/>
      <rect x="1128" y="308" width="264" height="61" fill="url(#selectedCell)"/>
      <path d="M468 154 C662 152 870 174 1186 336" fill="none" stroke="#14804e" stroke-width="5" stroke-linecap="round" opacity="0.72"/>
      <circle cx="468" cy="154" r="9" fill="#14804e"/>
      <circle cx="1186" cy="336" r="9" fill="#14804e"/>

      <text x="204" y="40" text-anchor="middle" fill="#71806b" font-size="20" font-weight="700">A</text>
      <text x="468" y="40" text-anchor="middle" fill="#71806b" font-size="20" font-weight="700">B</text>
      <text x="732" y="40" text-anchor="middle" fill="#71806b" font-size="20" font-weight="700">C</text>
      <text x="996" y="40" text-anchor="middle" fill="#71806b" font-size="20" font-weight="700">D</text>
      <text x="1260" y="40" text-anchor="middle" fill="#71806b" font-size="20" font-weight="700">E</text>

      <g fill="#71806b" font-size="17" font-weight="700">
        <text x="36" y="101" text-anchor="middle">1</text>
        <text x="36" y="162" text-anchor="middle">2</text>
        <text x="36" y="223" text-anchor="middle">3</text>
        <text x="36" y="284" text-anchor="middle">4</text>
        <text x="36" y="345" text-anchor="middle">5</text>
        <text x="36" y="406" text-anchor="middle">6</text>
        <text x="36" y="467" text-anchor="middle">7</text>
        <text x="36" y="528" text-anchor="middle">8</text>
        <text x="36" y="589" text-anchor="middle">9</text>
      </g>

      <g class="mono" font-size="26">
        <text x="570" y="163" text-anchor="end" fill="#12824d" font-weight="600">32</text>
        <text x="1098" y="163" text-anchor="end" fill="#242822">1200</text>
        <text x="1362" y="163" text-anchor="end" fill="#12824d" font-weight="600">38,400</text>
        <text x="570" y="224" text-anchor="end" fill="#242822">30</text>
        <text x="1098" y="224" text-anchor="end" fill="#242822">250</text>
        <text x="1362" y="224" text-anchor="end" fill="#242822">7,500</text>
        <text x="570" y="285" text-anchor="end" fill="#242822">18</text>
        <text x="1098" y="285" text-anchor="end" fill="#242822">300</text>
        <text x="1362" y="346" text-anchor="end" fill="#12824d" font-size="32" font-weight="600">51,300</text>
      </g>
    </g>
  </g>
</svg>`
}

async function renderPng(svg: string): Promise<Buffer> {
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-hero-workbook-api-'))
  const svgPath = join(tempRoot, 'hero.svg')

  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', ['--format=png', '--width', String(heroWidth), '--height', String(heroHeight), svgPath])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

const svg = await buildSvg()
const image = await renderPng(svg)
requirePngDimensions(image, 'rendered hero asset')

if (checkMode) {
  const existingSvg = await readFile(svgOutputPath, 'utf8')
  if (existingSvg !== svg) {
    throw new Error(`${svgOutputPath} is stale. Run pnpm docs:hero-asset:generate.`)
  }
  const existingImage = await readFile(outputPath)
  requirePngDimensions(existingImage, outputPath)
  console.log(`hero asset is current: ${outputPath}`)
} else {
  await mkdir(assetsRoot, { recursive: true })
  await writeFile(svgOutputPath, svg)
  await writeFile(outputPath, image)
  console.log(`wrote ${svgOutputPath}`)
  console.log(`wrote ${outputPath}`)
}
