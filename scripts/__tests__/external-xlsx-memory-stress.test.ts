import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  buildExternalXlsxStressPlan,
  externalXlsxStressSources,
  validateExternalXlsxStressPlan,
  type ExternalXlsxStressPlan,
} from '../external-xlsx-memory-stress.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'

describe('external XLSX memory stress plan', () => {
  it('tracks public Microsoft and Power BI workbooks including 100 MiB+ stress files', () => {
    const plan = buildExternalXlsxStressPlan({
      cacheDir: '/repo/.cache/external-xlsx-stress',
      maxRssBytes: 512 * 1024 * 1024,
    })

    expect(validateExternalXlsxStressPlan(plan)).toEqual([])
    expect(plan.sourceCount).toBe(externalXlsxStressSources.length)
    expect(plan.workbookCount).toBeGreaterThanOrEqual(7)
    expect(plan.giantWorkbookCount).toBeGreaterThanOrEqual(2)
    expect(plan.workbooks.map((workbook) => workbook.id)).toEqual(
      expect.arrayContaining(['powerpivot-tutorial-sample', 'contoso-sample-dax-formulas', 'powerbi-retail-analysis']),
    )
    expect(plan.sources.some((source) => source.downloadUrl.includes('download.microsoft.com'))).toBe(true)
    expect(plan.sources.some((source) => source.downloadUrl.includes('raw.githubusercontent.com'))).toBe(true)
  })

  it('rejects plans that do not include giant public workbook stress targets', () => {
    const validPlan = buildExternalXlsxStressPlan({ cacheDir: '/repo/.cache/external-xlsx-stress' })
    const invalidPlan: ExternalXlsxStressPlan = {
      ...validPlan,
      giantWorkbookCount: 0,
      workbooks: validPlan.workbooks.map((workbook) => Object.assign({}, workbook, { expectedMinBytes: 1024 })),
    }

    expect(validateExternalXlsxStressPlan(invalidPlan)).toEqual(
      expect.arrayContaining(['plan must include at least two 100 MiB+ workbook stress targets']),
    )
  })

  it('exposes package scripts for planning and running the external stress harness', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(packageJsonPath(), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])

    expect(scripts['external-xlsx-memory-stress:plan']).toBe('bun scripts/external-xlsx-memory-stress.ts plan')
    expect(scripts['external-xlsx-memory-stress']).toBe('bun scripts/external-xlsx-memory-stress.ts run')
  })
})

function packageJsonPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'package.json')
}
