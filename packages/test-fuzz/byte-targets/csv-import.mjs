import { importCsv } from '../../excel-import/dist/csv-import.js'

const decoder = new TextDecoder('utf-8', { fatal: false })

export function fuzz(data) {
  const text = decoder.decode(data)
  importCsv(text, 'fuzz.csv')
}
