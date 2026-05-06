import { describe, expect, it } from 'vitest'

import { parseCaptureArgs, parseEmitXlsxArgs } from '../capture-ui-responsiveness-same-corpus.ts'

describe('same-corpus UI responsiveness capture CLI', () => {
  it('builds a default Bilig benchmark URL from the selected corpus', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--corpus',
      'dense-mixed-250k',
    ])

    expect(args).toMatchObject({
      biligUrl: 'http://127.0.0.1:5173/?benchmarkCorpus=dense-mixed-250k',
      corpusId: 'dense-mixed-250k',
      deltaX: 0,
      deltaY: 720,
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      headless: true,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      sampleCount: 3,
    })
    expect(args.outputPath.endsWith('/tmp/ui-capture.json')).toBe(true)
  })

  it('accepts explicit browser and workload options', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--bilig-url',
      'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      '--samples',
      '5',
      '--delta-x',
      '1024',
      '--delta-y',
      '0',
      '--headed',
    ])

    expect(args).toMatchObject({
      biligUrl: 'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      deltaX: 1024,
      deltaY: 0,
      headless: false,
      sampleCount: 5,
    })
  })

  it('rejects missing incumbent URLs because the generated proof must be comparable', () => {
    expect(() => parseCaptureArgs(['--output', 'tmp/ui-capture.json'])).toThrow('Missing required arguments.')
  })

  it('parses XLSX emission mode for same-corpus setup', () => {
    const args = parseEmitXlsxArgs(['--emit-xlsx', 'tmp/ui-corpus', '--corpus', 'wide-mixed-variable-250k'])

    expect(args).toMatchObject({
      corpusId: 'wide-mixed-variable-250k',
    })
    expect(args?.targetDirectory.endsWith('/tmp/ui-corpus')).toBe(true)
  })

  it('rejects unknown corpus ids', () => {
    expect(() =>
      parseCaptureArgs([
        '--output',
        'tmp/ui-capture.json',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        '--microsoft-excel-web-url',
        'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
        '--corpus',
        'tiny-demo',
      ]),
    ).toThrow('Unexpected workbook benchmark corpus id: tiny-demo')
  })
})
