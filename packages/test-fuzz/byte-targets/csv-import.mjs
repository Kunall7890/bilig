import { importCsv } from '@bilig/excel-import'

const decoder = new TextDecoder('utf-8', { fatal: false })

export function fuzz(data) {
  const text = decoder.decode(data)
  importCsv(text, 'fuzz.csv')
}
