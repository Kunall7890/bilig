import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { forEachLargeSimpleInflatedZipEntryChunk } from '../xlsx-large-simple-stream-garbage.js'
import { readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple streamed ZIP garbage collection', () => {
  it('collects periodically while preserving streamed entry bytes', () => {
    const path = 'xl/worksheets/sheet1.xml'
    const text = 'a'.repeat(100)
    const zip = readXlsxZipEntriesLazy(zipSync({ [path]: strToU8(text) }))
    const chunks: Uint8Array[] = []
    let collections = 0

    expect(
      forEachLargeSimpleInflatedZipEntryChunk(zip, path, (chunk) => chunks.push(chunk), {
        chunkSize: 10,
        garbageCollectionIntervalBytes: 25,
        collectGarbage: () => {
          collections += 1
        },
      }),
    ).toBe(true)

    expect(Buffer.concat(chunks).toString()).toBe(text)
    expect(collections).toBe(3)
  })
})
