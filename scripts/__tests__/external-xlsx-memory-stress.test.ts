import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import type { ImportedWorkbook } from '../../packages/excel-import/src/workbook-import-result.js'
import { XLSX_CONTENT_TYPE } from '../../packages/excel-import/src/workbook-import-content-types.js'
import {
  buildExternalXlsxStressPlan,
  externalXlsxStressSources,
  validateExternalXlsxStressPlan,
  type ExternalXlsxStressPlan,
} from '../external-xlsx-memory-stress.ts'
import { shouldSummarizeFileBackedHeadlessInspect, summarizeExternalXlsxImportedWorkbook } from '../external-xlsx-memory-stress-worker.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'

describe('external XLSX memory stress plan', () => {
  it('tracks public Microsoft and Power BI workbooks including 100 MiB+ stress files', () => {
    const plan = buildExternalXlsxStressPlan({
      cacheDir: '/repo/.cache/external-xlsx-stress',
      maxRssBytes: 512 * 1024 * 1024,
    })

    expect(validateExternalXlsxStressPlan(plan)).toEqual([])
    expect(plan.sourceCount).toBe(externalXlsxStressSources.length)
    expect(plan.workbookCount).toBeGreaterThanOrEqual(19)
    expect(plan.giantWorkbookCount).toBeGreaterThanOrEqual(2)
    expect(plan.cellHeavyWorkbookCount).toBeGreaterThanOrEqual(3)
    expect(plan.workbooks.map((workbook) => workbook.id)).toEqual(
      expect.arrayContaining([
        'ons-cpi-mm23-current',
        'ons-trade-imports-current',
        'ons-life-expectancy-pivot',
        'govinfo-fy2027-outlays',
        'powerpivot-tutorial-sample',
        'contoso-sample-dax-formulas',
        'powerbi-adventureworks-sales',
        'powerbi-procurement-analysis',
        'powerbi-retail-analysis',
      ]),
    )
    expect(plan.sources.some((source) => source.downloadUrl.includes('download.microsoft.com'))).toBe(true)
    expect(plan.sources.some((source) => source.downloadUrl.includes('ons.gov.uk'))).toBe(true)
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

  it('rejects plans that do not include visible-cell-heavy public workbook stress targets', () => {
    const validPlan = buildExternalXlsxStressPlan({ cacheDir: '/repo/.cache/external-xlsx-stress' })
    const invalidPlan: ExternalXlsxStressPlan = {
      ...validPlan,
      cellHeavyWorkbookCount: 0,
      workbooks: validPlan.workbooks.map((workbook) => {
        const copy = Object.assign({}, workbook)
        delete copy.expectedMinCells
        return copy
      }),
    }

    expect(validateExternalXlsxStressPlan(invalidPlan)).toEqual(
      expect.arrayContaining(['plan must include at least three 1M+ visible-cell workbook stress targets']),
    )
  })

  it('exposes package scripts for planning and running the external stress harness', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(packageJsonPath(), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])

    expect(scripts['external-xlsx-memory-stress:plan']).toBe('bun scripts/external-xlsx-memory-stress.ts plan')
    expect(scripts['external-xlsx-memory-stress']).toBe('bun scripts/external-xlsx-memory-stress.ts run')
  })

  it('summarizes large-simple imports from stats without materializing lazy cells', () => {
    const throwingCellsTarget: ImportedWorkbook['snapshot']['sheets'][number]['cells'] = []
    const throwingCells = new Proxy(throwingCellsTarget, {
      get: (target, property, receiver) => {
        if (property === 'length') {
          return 2_450_603
        }
        if (property === Symbol.iterator) {
          throw new Error('lazy cells should not be iterated for stats-backed summaries')
        }
        return Reflect.get(target, property, receiver)
      },
    })
    const imported: ImportedWorkbook = {
      snapshot: {
        version: 1,
        workbook: {
          name: 'external-large.xlsx',
          metadata: { tables: [] },
        },
        sheets: [
          {
            id: 1,
            name: 'Sales',
            order: 0,
            metadata: { columns: [] },
            cells: throwingCells,
          },
        ],
      },
      workbookName: 'external-large.xlsx',
      sheetNames: ['Sales'],
      warnings: ['sample warning'],
      preview: {
        fileName: 'external-large.xlsx',
        contentType: XLSX_CONTENT_TYPE,
        fileSizeBytes: 1024,
        workbookName: 'external-large.xlsx',
        sheetCount: 0,
        sheets: [],
        warnings: [],
      },
      stats: {
        sheetCount: 1,
        cellCount: 2_450_603,
        formulaCellCount: 45_224,
        valueCellCount: 2_405_379,
        definedNameCount: 0,
        tableCount: 0,
        mergeCount: 0,
        conditionalFormatCount: 0,
        dataValidationCount: 0,
        warningCount: 1,
        dimensions: [],
        phaseTelemetry: [],
      },
    }

    expect(summarizeExternalXlsxImportedWorkbook(imported)).toMatchObject({
      importMode: 'public-snapshot',
      sheets: 1,
      cells: 2_450_603,
      formulas: 45_224,
      warnings: 1,
      workbookMetadataKeys: ['tables'],
      sheetMetadataKeys: ['columns'],
    })
  })

  it('uses the file-backed headless summary for formula-free large-simple workbooks even below the cell-heavy threshold', () => {
    expect(
      shouldSummarizeFileBackedHeadlessInspect(
        {
          workbookName: 'metadata-heavy.xlsx',
          sheetNames: ['Data'],
          warnings: ['External pivot caches were detected but not semantically imported during XLSX import.'],
          workbookMetadataKeys: ['pivotArtifacts'],
          sheetMetadataKeys: ['printPageSetup'],
          stats: {
            sheetCount: 1,
            cellCount: 37_062,
            formulaCellCount: 0,
            valueCellCount: 37_062,
            definedNameCount: 0,
            tableCount: 0,
            mergeCount: 0,
            conditionalFormatCount: 0,
            dataValidationCount: 0,
            warningCount: 1,
            dimensions: [],
            phaseTelemetry: [],
          },
        },
        512 * 1024,
      ),
    ).toBe(true)
    expect(
      shouldSummarizeFileBackedHeadlessInspect({
        workbookName: 'formula-small.xlsx',
        sheetNames: ['Data'],
        warnings: [],
        workbookMetadataKeys: [],
        sheetMetadataKeys: [],
        stats: {
          sheetCount: 1,
          cellCount: 37_062,
          formulaCellCount: 12,
          valueCellCount: 37_062,
          definedNameCount: 0,
          tableCount: 0,
          mergeCount: 0,
          conditionalFormatCount: 0,
          dataValidationCount: 0,
          warningCount: 0,
          dimensions: [],
          phaseTelemetry: [],
        },
      }),
    ).toBe(false)
    expect(
      shouldSummarizeFileBackedHeadlessInspect(
        {
          workbookName: 'tiny-complex.xlsx',
          sheetNames: ['Data'],
          warnings: ['External pivot caches were detected but not semantically imported during XLSX import.'],
          workbookMetadataKeys: ['pivotArtifacts'],
          sheetMetadataKeys: [],
          stats: {
            sheetCount: 1,
            cellCount: 597,
            formulaCellCount: 0,
            valueCellCount: 597,
            definedNameCount: 0,
            tableCount: 0,
            mergeCount: 0,
            conditionalFormatCount: 0,
            dataValidationCount: 0,
            warningCount: 1,
            dimensions: [],
            phaseTelemetry: [],
          },
        },
        5 * 1024 * 1024,
      ),
    ).toBe(false)
    expect(
      shouldSummarizeFileBackedHeadlessInspect(
        {
          workbookName: 'tiny-metadata-rich.xlsx',
          sheetNames: ['Data'],
          warnings: [],
          workbookMetadataKeys: ['dataModelArtifacts', 'styles'],
          sheetMetadataKeys: ['drawingArtifacts', 'hyperlinks', 'printerSettings', 'richTextArtifacts', 'styleRanges'],
          stats: {
            sheetCount: 1,
            cellCount: 12,
            formulaCellCount: 0,
            valueCellCount: 12,
            definedNameCount: 0,
            tableCount: 0,
            mergeCount: 0,
            conditionalFormatCount: 0,
            dataValidationCount: 0,
            warningCount: 0,
            dimensions: [],
            phaseTelemetry: [],
          },
        },
        2 * 1024 * 1024,
      ),
    ).toBe(true)
    expect(
      shouldSummarizeFileBackedHeadlessInspect(
        {
          workbookName: 'tiny-conditional-format-artifacts.xlsx',
          sheetNames: ['Data'],
          warnings: [],
          workbookMetadataKeys: [],
          sheetMetadataKeys: ['conditionalFormats', 'drawingArtifacts', 'printerSettings', 'richTextArtifacts', 'styleRanges'],
          stats: {
            sheetCount: 1,
            cellCount: 12,
            formulaCellCount: 0,
            valueCellCount: 12,
            definedNameCount: 0,
            tableCount: 0,
            mergeCount: 0,
            conditionalFormatCount: 1,
            dataValidationCount: 0,
            warningCount: 0,
            dimensions: [],
            phaseTelemetry: [],
          },
        },
        2 * 1024 * 1024,
      ),
    ).toBe(false)
    expect(
      shouldSummarizeFileBackedHeadlessInspect(
        {
          workbookName: 'giant-package.xlsx',
          sheetNames: ['Data'],
          warnings: ['External pivot caches were detected but not semantically imported during XLSX import.'],
          workbookMetadataKeys: ['pivotArtifacts'],
          sheetMetadataKeys: [],
          stats: {
            sheetCount: 1,
            cellCount: 735,
            formulaCellCount: 0,
            valueCellCount: 735,
            definedNameCount: 0,
            tableCount: 0,
            mergeCount: 0,
            conditionalFormatCount: 0,
            dataValidationCount: 0,
            warningCount: 1,
            dimensions: [],
            phaseTelemetry: [],
          },
        },
        221 * 1024 * 1024,
      ),
    ).toBe(true)
  })
})

function packageJsonPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'package.json')
}
