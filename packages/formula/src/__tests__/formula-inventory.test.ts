import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { formulaInventory, formulaInventorySummary } from '../generated/formula-inventory.js'
import { normalizeFormulaName } from '../builtin-capabilities.js'

interface FormulaInventorySource {
  version: number
  entries: Array<{ name: string; odfStatus: string; inOfficeList: boolean }>
}

describe('formula inventory', () => {
  it('tracks the canonical unified formula count', () => {
    expect(formulaInventorySummary.total).toBe(formulaInventory.length)
    expect(formulaInventory.length).toBeGreaterThan(500)
  })

  it('keeps the generated inventory aligned with the canonical source inventory', () => {
    const source = readFormulaInventorySource()
    expect(source.version).toBe(1)
    expect(formulaInventory).toHaveLength(source.entries.length)
    expect(formulaInventory.map((entry) => entry.name)).toEqual(source.entries.map((entry) => normalizeFormulaName(entry.name)))
  })

  it('tracks summary counts consistently with the generated inventory', () => {
    const registeredInCodebase = formulaInventory.filter((entry) => entry.registeredInCodebase)
    const placeholders = formulaInventory.filter((entry) => entry.placeholder)

    expect(formulaInventorySummary.registeredInCodebase).toBe(registeredInCodebase.length)
    expect(formulaInventorySummary.missingInCodebase).toBe(formulaInventory.length - registeredInCodebase.length)
    expect(formulaInventorySummary.placeholders).toBe(placeholders.length)
  })

  it('keeps runtime and protocol reporting for key formulas', () => {
    const letEntry = formulaInventory.find((entry) => entry.name === 'LET')
    const sumEntry = formulaInventory.find((entry) => entry.name === 'SUM')
    const copilotEntry = formulaInventory.find((entry) => entry.name === 'COPILOT')
    const imageEntry = formulaInventory.find((entry) => entry.name === 'IMAGE')
    const importRangeEntry = formulaInventory.find((entry) => entry.name === 'IMPORTRANGE')
    const pyEntry = formulaInventory.find((entry) => entry.name === 'PY')
    const iserrorEntry = formulaInventory.find((entry) => entry.name === 'ISERROR')

    expect(letEntry).toMatchObject({
      registeredInCodebase: true,
      protocolId: expect.any(Number),
      protocolSupportsWasm: false,
      deterministic: 'deterministic',
      jsStatus: 'special-js-only',
      wasmStatus: 'not-started',
    })
    expect(sumEntry).toMatchObject({
      registeredInCodebase: true,
      protocolSupportsWasm: true,
      runtimeStatus: 'implemented',
    })
    expect(iserrorEntry).toMatchObject({
      registeredInCodebase: true,
      protocolId: expect.any(Number),
      protocolSupportsWasm: false,
      runtimeStatus: 'implemented',
      wasmStatus: 'not-started',
    })
    expect(copilotEntry).toMatchObject({
      inOfficeList: true,
      deterministic: 'provider-backed',
      protocolId: undefined,
      runtimeStatus: 'implemented',
      placeholder: false,
      registeredInCodebase: true,
    })
    expect(imageEntry).toMatchObject({
      deterministic: 'provider-backed',
      protocolId: undefined,
    })
    expect(importRangeEntry).toMatchObject({
      inOfficeList: false,
      deterministic: 'provider-backed',
      protocolId: undefined,
      runtimeStatus: 'implemented',
      placeholder: false,
      registeredInCodebase: true,
    })
    expect(pyEntry).toMatchObject({
      inOfficeList: true,
      deterministic: 'provider-backed',
      protocolId: undefined,
      runtimeStatus: 'implemented',
      placeholder: false,
      registeredInCodebase: true,
    })
  })

  it('keeps current Microsoft-listed compatibility functions in the Office inventory', () => {
    const officeListedNames = new Set(formulaInventory.filter((entry) => entry.inOfficeList).map((entry) => entry.name))

    for (const name of ['COPILOT', 'FORECAST', 'ISODD', 'ISREF', 'JIS']) {
      expect(officeListedNames.has(name)).toBe(true)
    }
  })

  it('keeps every deterministic runtime formula visible in the protocol inventory', () => {
    const missingProtocolEntries = formulaInventory.filter(
      (entry) => entry.deterministic === 'deterministic' && entry.registeredInCodebase && entry.protocolId === undefined,
    )

    expect(formulaInventorySummary.runtimeRegisteredMissingProtocol).toBe(0)
    expect(missingProtocolEntries).toEqual([])
  })

  it('counts runtime builtins registered by source factories', () => {
    const factoryRegisteredFunctions = [
      'ACCRINT',
      'AMORLINC',
      'ARABIC',
      'AVEDEV',
      'COMPLEX',
      'CONCATENATE',
      'DATEVALUE',
      'ENCODEURL',
      'IMABS',
      'LEGACY.NORMSDIST',
      'LEGACY.NORMSINV',
      'MDETERM',
      'MMULT',
      'MULTINOMIAL',
      'PERCENTOF',
      'PROPER',
      'REGEXMATCH',
      'REGEXTEST',
      'ROMAN',
      'SKEWP',
      'COUNTUNIQUEIFS',
    ]

    for (const name of factoryRegisteredFunctions) {
      expect(formulaInventory.find((entry) => entry.name === name)).toMatchObject({
        name,
        runtimeStatus: 'implemented',
        registeredInCodebase: true,
        placeholder: false,
      })
    }
  })
})

function readFormulaInventorySource(): FormulaInventorySource {
  const parsed = JSON.parse(readFileSync(new URL('../formula-inventory-source.json', import.meta.url), 'utf8')) as unknown
  if (!isFormulaInventorySource(parsed)) {
    throw new Error('Invalid formula inventory source fixture')
  }
  return parsed
}

function isFormulaInventorySource(value: unknown): value is FormulaInventorySource {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const candidate = value as { version?: unknown; entries?: unknown }
  return (
    candidate.version === 1 &&
    Array.isArray(candidate.entries) &&
    candidate.entries.every(
      (entry) =>
        typeof entry === 'object' &&
        entry !== null &&
        typeof Reflect.get(entry, 'name') === 'string' &&
        typeof Reflect.get(entry, 'odfStatus') === 'string' &&
        typeof Reflect.get(entry, 'inOfficeList') === 'boolean',
    )
  )
}
