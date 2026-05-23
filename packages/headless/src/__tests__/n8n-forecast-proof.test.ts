import { describe, expect, it } from 'vitest'

import { createN8nForecastProof } from '../n8n-forecast-proof.js'
import { n8nForecastServerHelpText, parseN8nForecastServerCliArgs } from '../n8n-forecast-server-cli.js'

describe('n8n formula readback proof', () => {
  it('returns formula readback and restore proof for one input edit', () => {
    expect(
      createN8nForecastProof({
        address: 'B3',
        value: 0.4,
      }),
    ).toMatchObject({
      verified: true,
      editedCell: 'Inputs!B3',
      before: {
        expectedCustomers: 5,
        expectedArr: 60000,
        expansionArr: 66000,
        targetGap: -34000,
      },
      after: {
        expectedCustomers: 8,
        expectedArr: 96000,
        expansionArr: 105600,
        targetGap: 5600,
      },
      formulaContracts: {
        expectedCustomers: '=Inputs!B2*Inputs!B3',
        expectedArr: '=B2*Inputs!B4',
        expansionArr: '=B3*Inputs!B5',
        targetGap: '=B4-100000',
      },
      checks: {
        previousValue: 0.25,
        newValue: 0.4,
        formulasPersisted: true,
        restoredMatchesAfter: true,
        computedOutputChanged: true,
      },
    })
  })

  it('rejects edits outside the demo input cells', () => {
    expect(() =>
      createN8nForecastProof({
        address: 'C9',
        value: 0.4,
      }),
    ).toThrow('Editable input address must be one of B2, B3, B4, B5')
  })

  it('parses the local server CLI options', () => {
    expect(parseN8nForecastServerCliArgs([], {})).toEqual({
      help: false,
      host: '127.0.0.1',
      port: 4321,
    })
    expect(
      parseN8nForecastServerCliArgs(['--host', '0.0.0.0', '--port', '8787'], {
        BILIG_N8N_PORT: '9999',
      }),
    ).toEqual({
      help: false,
      host: '0.0.0.0',
      port: 8787,
    })
    expect(n8nForecastServerHelpText()).toContain('bilig-n8n-formula-server')
    expect(() => parseN8nForecastServerCliArgs(['--bad'], {})).toThrow(
      'Unknown bilig-n8n-formula-server argument',
    )
  })
})
