import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'
import { createRegionGraph } from '../deps/region-graph.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('RegionGraph', () => {
  it('interns single-column regions canonically and collects point dependents through interval ownership', () => {
    const workbook = new WorkbookStore('region-graph')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const same = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const overlapping = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 2,
      rowEnd: 4,
      col: 0,
    })

    expect(same).toBe(first)
    expect(regionGraph.hasFormulaSubscriptions()).toBe(false)

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [overlapping])
    expect(regionGraph.hasFormulaSubscriptions()).toBe(true)
    expect(regionGraph.getFormulaSubscriptionCount()).toBe(2)
    expect(regionGraph.hasFormulaSubscriptionsOverlappingRange(sheet.id, 0, 4, 0, 0)).toBe(true)
    expect(regionGraph.hasFormulaSubscriptionsOverlappingRange(sheet.id, 5, 8, 0, 0)).toBe(false)
    expect(regionGraph.hasFormulaSubscriptionsOverlappingRange(sheet.id, 0, 4, 1, 1)).toBe(false)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 0, 0)).toBe(10)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 2, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 2, 0)).toBe(-2)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 4, 0)).toBe(20)
    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([first])

    regionGraph.clearFormulaSubscriptions(10)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 2, 0)]).toEqual([20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 2, 0)).toBe(20)
    regionGraph.clearFormulaSubscriptions(20)
    expect(regionGraph.hasFormulaSubscriptions()).toBe(false)
    expect(regionGraph.getFormulaSubscriptionCount()).toBe(0)
  })

  it('can prepare query indices eagerly before the first point lookup', () => {
    const workbook = new WorkbookStore('region-graph-prepare')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 16,
      rowEnd: 47,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])
    regionGraph.prepareQueryIndices()

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])

    regionGraph.clearFormulaSubscriptions(10)
    regionGraph.prepareQueryIndices()

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)]).toEqual([20])
  })

  it('answers the first dirty point query without building an interval tree', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-dirty-point-query')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 16,
      rowEnd: 47,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])
    expect(counters.regionQueryIndexBuilds).toBe(0)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect(counters.regionQueryIndexBuilds).toBe(0)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 47, 0)]).toEqual([20])
    expect(counters.regionQueryIndexBuilds).toBe(1)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 32, 0)]).toEqual([20])
    expect(counters.regionQueryIndexBuilds).toBe(1)
  })

  it('answers dirty point queries outside subscribed row bounds without building interval trees', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-out-of-bounds-query')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 64,
      rowEnd: 96,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 256, 0)]).toEqual([])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 256, 0)).toBe(-1)
    expect(counters.regionQueryIndexBuilds).toBe(0)

    regionGraph.clearFormulaSubscriptions(20)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 64, 0)]).toEqual([])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 64, 0)).toBe(-1)
    expect(counters.regionQueryIndexBuilds).toBe(0)
  })

  it('checks rectangular subscription overlap from maintained column row bounds', () => {
    const workbook = new WorkbookStore('region-graph-rect-overlap')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 64,
      rowEnd: 96,
      col: 2,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])

    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 32, 63, 0, 2)).toBe(false)
    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 24, 40, 0, 1)).toBe(true)
    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 40, 80, 1, 2)).toBe(true)
    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 0, 100, 3, 4)).toBe(false)

    regionGraph.clearFormulaSubscriptions(20)
    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 64, 96, 2, 2)).toBe(false)
  })

  it('replaces a single formula region subscription without disturbing other subscribers', () => {
    const workbook = new WorkbookStore('region-graph-replace-single')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 5,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [first])
    regionGraph.replaceSingleFormulaSubscription(10, first, second)

    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([second])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 1, 0)]).toEqual([20])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([10])
  })

  it('replaces many single formula region subscriptions in one column batch', () => {
    const workbook = new WorkbookStore('region-graph-replace-single-batch')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 5,
      col: 0,
    })
    const third = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 6,
      rowEnd: 8,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])
    regionGraph.replaceFormulaSubscriptions(30, [third])
    regionGraph.replaceSingleFormulaSubscriptions([
      { formulaCellIndex: 10, previousRegionId: first, nextRegionId: second },
      { formulaCellIndex: 20, previousRegionId: second, nextRegionId: third },
    ])

    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([second])
    expect(regionGraph.getFormulaSubscriptions(20)).toEqual([third])
    expect(regionGraph.getFormulaSubscriptions(30)).toEqual([third])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 1, 0)]).toEqual([])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([10])
    expect(new Set(regionGraph.collectFormulaDependentsForCell(sheet.id, 7, 0))).toEqual(new Set([20, 30]))
  })

  it('deduplicates dependents when one formula subscribes to multiple matching regions', () => {
    const workbook = new WorkbookStore('region-graph-dedupe')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 5,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 2,
      rowEnd: 7,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first, second])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 3, 0)]).toEqual([10])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 3, 0)).toBe(10)
  })

  it('deduplicates large overlapping range fanouts without dropping dependents', () => {
    const workbook = new WorkbookStore('region-graph-large-fanout-dedupe')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const expected = new Set<number>()
    for (let index = 0; index < 96; index += 1) {
      const region = regionGraph.internSingleColumnRegion({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: index + 1,
        col: 0,
      })
      const formulaCellIndex = 1_000 + index
      regionGraph.replaceFormulaSubscriptions(formulaCellIndex, [region])
      expected.add(formulaCellIndex)
    }
    const duplicateFirst = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 12,
      col: 0,
    })
    const duplicateSecond = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 24,
      col: 0,
    })
    regionGraph.replaceFormulaSubscriptions(5_000, [duplicateFirst, duplicateSecond])
    expected.add(5_000)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)].toSorted((left, right) => left - right)).toEqual(
      [...expected].toSorted((left, right) => left - right),
    )
  })

  it('keeps large sliding-window first point lookups on the dirty ordered scan path', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-sliding-large')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    for (let row = 0; row < 1_500; row += 1) {
      const region = regionGraph.internSingleColumnRegion({
        sheetName: 'Sheet1',
        rowStart: row,
        rowEnd: row + 31,
        col: 0,
      })
      regionGraph.replaceFormulaSubscriptions(10_000 + row, [region])
    }

    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 0, 0)).toBe(10_000)
    expect(counters.regionQueryIndexBuilds).toBe(0)
  })

  it('handles empty, duplicate, mismatched, and reset subscription operations', () => {
    const workbook = new WorkbookStore('region-graph-replacement-edges')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 5,
      col: 0,
    })
    const third = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 6,
      rowEnd: 8,
      col: 0,
    })

    expect(regionGraph.getRegion(first)).toMatchObject({ sheetId: sheet.id, rowStart: 0, rowEnd: 2, col: 0 })
    expect(() => regionGraph.internSingleColumnRegion({ sheetName: 'Missing', rowStart: 0, rowEnd: 1, col: 0 })).toThrow(
      'Unknown sheet for region graph: Missing',
    )
    expect(regionGraph.getFormulaSubscriptions(1234)).toEqual([])
    expect(regionGraph.hasFormulaSubscriptionsOverlappingRange(sheet.id, 4, 2, 0, 0)).toBe(false)
    expect(regionGraph.hasFormulaSubscriptionsIntersectingRect(sheet.id, 4, 2, 0, 0)).toBe(false)

    regionGraph.replaceFormulaSubscriptions(10, [first, first, second])
    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([first, second])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 7, 0)]).toEqual([])

    regionGraph.replaceSingleFormulaSubscription(20, first, second)
    expect(regionGraph.getFormulaSubscriptions(20)).toEqual([second])

    regionGraph.replaceFormulaSubscriptions(30, [first, first])
    expect(regionGraph.getFormulaSubscriptions(30)).toEqual([first])
    regionGraph.replaceSingleFormulaSubscription(30, first, third)
    expect(regionGraph.getFormulaSubscriptions(30)).toEqual([third])

    regionGraph.replaceSingleFormulaSubscriptions([])
    regionGraph.replaceSingleFormulaSubscriptions([
      { formulaCellIndex: 40, previousRegionId: first, nextRegionId: first },
      { formulaCellIndex: 50, previousRegionId: first, nextRegionId: second },
    ])
    expect(regionGraph.getFormulaSubscriptions(40)).toEqual([])
    expect(regionGraph.getFormulaSubscriptions(50)).toEqual([second])

    regionGraph.reset()
    expect(regionGraph.hasFormulaSubscriptions()).toBe(false)
    expect(regionGraph.hasFormulaSubscriptionsForColumn(sheet.id, 0)).toBe(false)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([])
  })

  it('falls back to interval-tree single-dependent queries when point impact indexing is disabled', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-large-interval-single')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    const left = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 300,
      col: 0,
    })
    const right = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 400,
      rowEnd: 500,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [left])
    regionGraph.replaceFormulaSubscriptions(20, [right])

    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 10, 0)).toBe(10)
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 450, 0)).toBe(20)
    expect(counters.regionQueryIndexBuilds).toBe(0)

    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 10, 0)).toBe(10)
    expect(counters.regionQueryIndexBuilds).toBe(1)
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 450, 0)).toBe(20)
    expect(counters.regionQueryIndexBuilds).toBe(1)
  })
})
