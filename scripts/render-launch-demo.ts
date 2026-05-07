import { chromium } from '@playwright/test'
import { Buffer } from 'node:buffer'
import { copyFile, mkdir, mkdtemp, readFile, stat, unlink } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const demoWidth = 1280
const demoHeight = 720
const demoDurationMs = 9_000
const minDemoBytes = 80_000
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const outputPath = join(repoRoot, 'docs', 'assets', 'product-hunt-demo.webm')
const checkMode = process.argv.includes('--check')

const html = String.raw`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <style>
      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
        overflow: hidden;
        background:
          linear-gradient(rgba(15, 23, 42, 0.055) 1px, transparent 1px),
          linear-gradient(90deg, rgba(15, 23, 42, 0.055) 1px, transparent 1px),
          #eef4f7;
        background-size: 34px 34px;
        color: #101820;
        font-family: Inter, Arial, Helvetica, sans-serif;
      }

      .frame {
        display: grid;
        width: ${demoWidth}px;
        height: ${demoHeight}px;
        grid-template-columns: 0.9fr 1.1fr;
        gap: 30px;
        padding: 44px;
      }

      .left,
      .right {
        min-width: 0;
        border: 1px solid #c8d5de;
        border-radius: 18px;
        background: #ffffff;
        box-shadow: 0 24px 70px rgba(15, 23, 42, 0.13);
      }

      .left {
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        padding: 34px 38px;
      }

      .brand {
        display: flex;
        align-items: center;
        gap: 16px;
        color: #2f6f47;
        font-size: 21px;
        font-weight: 780;
      }

      .mark {
        display: grid;
        width: 50px;
        height: 50px;
        grid-template-columns: repeat(2, 1fr);
        grid-template-rows: repeat(2, 1fr);
        gap: 7px;
        border: 1px solid #a9bdc9;
        border-radius: 12px;
        background: #101820;
        padding: 10px;
      }

      .mark span {
        border-radius: 3px;
        background: #77c98b;
      }

      h1 {
        margin: 42px 0 0;
        color: #111820;
        font-size: 62px;
        font-weight: 820;
        letter-spacing: 0;
        line-height: 1.02;
      }

      .subtitle {
        margin: 22px 0 0;
        color: #526273;
        font-size: 27px;
        font-weight: 620;
        letter-spacing: 0;
        line-height: 1.3;
      }

      .url {
        color: #526273;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
        font-size: 21px;
        letter-spacing: 0;
      }

      .right {
        display: grid;
        grid-template-rows: 218px 250px 1fr;
        gap: 16px;
        padding: 24px;
      }

      .panel {
        overflow: hidden;
        border: 1px solid #cbd8e2;
        border-radius: 14px;
        background: #f8fbfd;
      }

      .panel-title {
        display: flex;
        min-height: 44px;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #d7e1ea;
        background: #ffffff;
        padding: 0 18px;
        color: #506173;
        font-size: 15px;
        font-weight: 800;
        text-transform: uppercase;
      }

      .code {
        background: #111820;
        color: #dce9f2;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
        font-size: 18px;
        line-height: 1.45;
        padding: 18px 22px;
      }

      .prompt,
      .ok {
        color: #7ed894;
      }

      .muted {
        color: #95a7b7;
      }

      .grid {
        width: 100%;
        border-collapse: collapse;
        font-size: 17px;
      }

      .grid th,
      .grid td {
        height: 39px;
        border-right: 1px solid #d7e1ea;
        border-bottom: 1px solid #d7e1ea;
        padding: 0 14px;
        text-align: left;
      }

      .grid th {
        background: #edf4f8;
        color: #617486;
        font-size: 13px;
        font-weight: 820;
        text-transform: uppercase;
      }

      .grid td {
        color: #253241;
        font-weight: 660;
      }

      .number {
        color: #207146;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
        font-weight: 760;
      }

      .highlight {
        animation: flash 9s ease-in-out forwards;
      }

      .cursor {
        display: inline-block;
        width: 11px;
        height: 20px;
        margin-left: 4px;
        background: #7ed894;
        animation: blink 900ms steps(2, start) infinite;
        vertical-align: -3px;
      }

      .phase {
        opacity: 0;
        transform: translateY(10px);
        animation: phase-in 9s ease-out forwards;
      }

      .phase.two {
        animation-delay: 2.4s;
      }

      .phase.three {
        animation-delay: 4.8s;
      }

      .phase.four {
        animation-delay: 6.7s;
      }

      @keyframes phase-in {
        0%,
        18% {
          opacity: 0;
          transform: translateY(10px);
        }
        25%,
        100% {
          opacity: 1;
          transform: translateY(0);
        }
      }

      @keyframes flash {
        0%,
        44% {
          background: transparent;
        }
        52%,
        66% {
          background: #e3f6e9;
        }
        100% {
          background: transparent;
        }
      }

      @keyframes blink {
        to {
          opacity: 0;
        }
      }
    </style>
  </head>
  <body>
    <main class="frame">
      <section class="left">
        <div>
          <div class="brand">
            <div class="mark" aria-hidden="true">
              <span></span>
              <span></span>
              <span></span>
              <span></span>
            </div>
            <span>@bilig/headless</span>
          </div>
          <h1>workbook edits with proof, not pixels</h1>
          <p class="subtitle">load data, write formulas, recalc, persist, restore, and read back exact values from node.</p>
        </div>
        <div class="url">github.com/proompteng/bilig</div>
      </section>
      <section class="right">
        <section class="panel">
          <div class="panel-title">
            <span>agent command</span>
            <span>node</span>
          </div>
          <div class="code">
            <div class="phase">const workbook = WorkPaper.buildFromSheets(data)</div>
            <div class="phase two">workbook.setCellContents(arrCell, "=B4*240*1.2*12")</div>
            <div class="phase three">const after = workbook.getCellValue(arrCell)<span class="cursor"></span></div>
          </div>
        </section>
        <section class="panel">
          <div class="panel-title">
            <span>computed workbook</span>
            <span>after edit</span>
          </div>
          <table class="grid">
            <thead>
              <tr>
                <th>metric</th>
                <th>value</th>
                <th>formula</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>visitors</td>
                <td class="number">650</td>
                <td>input</td>
              </tr>
              <tr>
                <td>conversion</td>
                <td class="number">10%</td>
                <td>input</td>
              </tr>
              <tr>
                <td>customers</td>
                <td class="number">65</td>
                <td>=B2*B3</td>
              </tr>
              <tr class="highlight">
                <td>annual arr</td>
                <td class="number">224640</td>
                <td>=B4*240*1.2*12</td>
              </tr>
            </tbody>
          </table>
        </section>
        <section class="panel">
          <div class="panel-title">
            <span>verification</span>
            <span>restored document</span>
          </div>
          <div class="code">
            <div class="phase three"><span class="muted">before:</span> arr = 172,800</div>
            <div class="phase three"><span class="muted">after:</span> arr = 224,640</div>
            <div class="phase four ok">ok formula persisted after restore</div>
          </div>
        </section>
      </section>
    </main>
  </body>
</html>`

async function requireCurrentDemoAsset(): Promise<void> {
  const info = await stat(outputPath)
  if (!info.isFile() || info.size < minDemoBytes) {
    throw new Error(`${outputPath} is missing or too small. Run pnpm docs:launch-demo:generate.`)
  }

  const header = await readFile(outputPath)
  if (!header.subarray(0, 4).equals(Buffer.from([0x1a, 0x45, 0xdf, 0xa3]))) {
    throw new Error(`${outputPath} is not a WebM/EBML file`)
  }
}

async function renderDemo(): Promise<void> {
  const videoDir = await mkdtemp(join(tmpdir(), 'bilig-launch-demo-'))
  const browser = await chromium.launch()
  try {
    const context = await browser.newContext({
      viewport: { width: demoWidth, height: demoHeight },
      colorScheme: 'light',
      recordVideo: {
        dir: videoDir,
        size: { width: demoWidth, height: demoHeight },
      },
    })
    const page = await context.newPage()
    await page.setContent(html, { waitUntil: 'load' })
    await page.waitForTimeout(demoDurationMs)
    const video = page.video()
    if (video === null) {
      throw new Error('Playwright did not create a launch demo video')
    }
    await context.close()

    await mkdir(dirname(outputPath), { recursive: true })
    const renderedPath = await video.path()
    await copyFile(renderedPath, outputPath)
    await unlink(renderedPath)
  } finally {
    await browser.close()
  }
}

if (checkMode) {
  await requireCurrentDemoAsset()
  console.log(`launch demo is current: ${outputPath}`)
} else {
  await renderDemo()
  await requireCurrentDemoAsset()
  console.log(`wrote ${outputPath}`)
}
