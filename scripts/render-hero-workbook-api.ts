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
      <stop offset="0" stop-color="#fbfbf6"/>
      <stop offset="1" stop-color="#edf5ec"/>
    </linearGradient>
    <linearGradient id="selectedCell" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#c7f3d7"/>
      <stop offset="1" stop-color="#aeecc9"/>
    </linearGradient>
    <linearGradient id="codeRail" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#102017"/>
      <stop offset="1" stop-color="#0d1513"/>
    </linearGradient>
    <filter id="softShadow" x="-10%" y="-10%" width="120%" height="120%">
      <feDropShadow dx="0" dy="14" stdDeviation="18" flood-color="#2b3328" flood-opacity="0.07"/>
    </filter>
  </defs>

  <g transform="translate(68 96)" filter="url(#softShadow)">
    <g transform="translate(44 26)">
      <rect x="0" y="0" width="920" height="62" rx="10" fill="#fbfbf7" stroke="#d4ddd1"/>
      <rect x="18" y="15" width="92" height="32" rx="6" fill="#eaf1e7" stroke="#d4ddd1"/>
      <text x="64" y="37" text-anchor="middle" fill="#5b6559" font-size="17" font-weight="700">D5</text>
      <text x="142" y="39" class="mono" fill="#14784b" font-size="25" font-weight="600">=SUM(D2:D4)</text>
    </g>

    <g transform="translate(44 136)">
      <rect x="0" y="0" width="920" height="462" fill="url(#sheet)" stroke="#cbd8c7" stroke-width="2"/>
      <rect x="0" y="0" width="920" height="58" fill="#e7efe4"/>
      <rect x="0" y="0" width="72" height="462" fill="#e7efe4"/>
      <path d="M72 0 V462 M304 0 V462 M536 0 V462 M688 0 V462 M920 0 V462" stroke="#d7dfd4"/>
      <path d="M0 58 H920 M0 126 H920 M0 194 H920 M0 262 H920 M0 330 H920 M0 398 H920" stroke="#d7dfd4"/>

      <rect x="304" y="126" width="232" height="68" fill="url(#selectedCell)"/>
      <rect x="688" y="330" width="232" height="68" fill="url(#selectedCell)"/>
      <path d="M420 160 C524 158 614 188 680 258 C714 302 720 340 720 364" fill="none" stroke="#14784b" stroke-width="6" stroke-linecap="round" opacity="0.7"/>
      <circle cx="420" cy="160" r="10" fill="#14784b"/>
      <circle cx="720" cy="364" r="10" fill="#14784b"/>

      <g fill="#72806d" font-size="18" font-weight="700">
        <text x="188" y="37" text-anchor="middle">A</text>
        <text x="420" y="37" text-anchor="middle">B</text>
        <text x="612" y="37" text-anchor="middle">C</text>
        <text x="804" y="37" text-anchor="middle">D</text>
        <text x="36" y="100" text-anchor="middle">1</text>
        <text x="36" y="168" text-anchor="middle">2</text>
        <text x="36" y="236" text-anchor="middle">3</text>
        <text x="36" y="304" text-anchor="middle">4</text>
        <text x="36" y="372" text-anchor="middle">5</text>
      </g>

      <g font-size="24" fill="#262b25">
        <text x="112" y="102">Region</text>
        <text x="460" y="102" text-anchor="end">Customers</text>
        <text x="648" y="102" text-anchor="end">ARPA</text>
        <text x="874" y="102" text-anchor="end">Revenue</text>
        <text x="112" y="170">West</text>
        <text x="460" y="170" text-anchor="end" fill="#14784b" font-weight="700">32</text>
        <text x="648" y="170" text-anchor="end">1200</text>
        <text x="874" y="170" text-anchor="end" fill="#14784b" font-weight="700">38,400</text>
        <text x="112" y="238">East</text>
        <text x="460" y="238" text-anchor="end">30</text>
        <text x="648" y="238" text-anchor="end">250</text>
        <text x="874" y="238" text-anchor="end">7,500</text>
        <text x="112" y="306">Central</text>
        <text x="460" y="306" text-anchor="end">18</text>
        <text x="648" y="306" text-anchor="end">300</text>
        <text x="874" y="306" text-anchor="end">5,400</text>
        <text x="112" y="374" font-weight="700">Total</text>
        <text x="874" y="374" text-anchor="end" fill="#14784b" font-size="31" font-weight="700">51,300</text>
      </g>
    </g>

    <g transform="translate(1032 162)">
      <rect x="0" y="0" width="280" height="58" rx="10" fill="#0e1913"/>
      <text x="24" y="37" fill="#dbe9df" font-size="20" font-weight="700">B2</text>
      <text x="72" y="37" fill="#91a397" font-size="20">20</text>
      <path d="M118 29 H172" stroke="#91a397" stroke-width="2"/>
      <path d="M172 29 l-12 -8 M172 29 l-12 8" stroke="#91a397" stroke-width="2" fill="none"/>
      <text x="206" y="37" fill="#35d179" font-size="20" font-weight="700">32</text>
    </g>

    <g transform="translate(1032 268)">
      <rect x="0" y="0" width="360" height="70" rx="10" fill="#0e1913"/>
      <text x="24" y="42" class="mono" fill="#35d179" font-size="24" font-weight="600">getCellValue(D5)</text>
    </g>

    <g transform="translate(1032 372)">
      <rect x="0" y="0" width="360" height="116" rx="10" fill="#0e1913"/>
      <text x="24" y="42" fill="#dbe9df" font-size="19">after restore</text>
      <text x="24" y="88" fill="#35d179" font-size="42" font-weight="700">51,300</text>
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
